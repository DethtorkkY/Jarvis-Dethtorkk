"""JARVIS: offline Windows voice assistant with game controls and profiles."""
from __future__ import annotations

import json
import logging
import queue
import re
import csv
import subprocess
import sys
import time
from difflib import SequenceMatcher
from pathlib import Path

import sounddevice as sd
import win32com.client
from vosk import KaldiRecognizer, Model, SetLogLevel

APP_DIR = Path(__file__).resolve().parent
MODEL_DIR = APP_DIR / "models" / "vosk-model-small-ru-0.22"
CONFIG_FILE = APP_DIR / "game_profiles.json"
LOG_FILE = APP_DIR / "jarvis.log"
WAKE_WORDS = ("джарвис", "джарвес", "jarvis")
GAME_WORDS = {"игра", "игру", "игры", "играть", "поиграть", "поиграем", "поиграй", "катка", "катку", "каточку", "гамать", "гамаем", "гейм", "гейминг", "game", "gaming"}
PROFILE_ALIASES = {
    "competitive": {"компетитив", "соревновательный", "соревновательный", "рейтинг", "ранкед", "ranked"},
    "streaming": {"стрим", "стриминг", "stream", "streaming"},
    "chill": {"чил", "чилл", "отдых", "спокойный", "chill"},
    "afk": {"афк", "afk", "отошел", "отошла"},
}

logging.basicConfig(filename=LOG_FILE, level=logging.INFO,
                    format="%(asctime)s %(levelname)s %(message)s", encoding="utf-8")


def normalize(text: str) -> str:
    return re.sub(r"[^a-zа-яё0-9 ]", " ", text.lower()).replace("ё", "е")


def load_config() -> dict:
    with CONFIG_FILE.open(encoding="utf-8") as file:
        return json.load(file)


def create_speaker():
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    speaker.Volume = 100
    return speaker


def say(speaker, message: str) -> None:
    logging.info("Speaking: %s", message)
    speaker.Speak(message)


def game_label(game: dict) -> str:
    return game.get("label") or game.get("aliases", ["игру"])[0].capitalize()


def steam_games(config: dict) -> dict[str, dict]:
    """Read installed Steam games from appmanifest files, without hardcoding IDs."""
    steam = Path(config["games"]["steam"]["launcher"])
    manifests = steam.parent / "steamapps"
    found: dict[str, dict] = {}
    for manifest in manifests.glob("appmanifest_*.acf"):
        content = manifest.read_text(encoding="utf-8", errors="ignore")
        appid = re.search(r'"appid"\s+"(\d+)"', content)
        name = re.search(r'"name"\s+"([^"]+)"', content)
        if not appid or not name:
            continue
        key = f"steam_{appid.group(1)}"
        found[key] = {
            "label": name.group(1), "aliases": [normalize(name.group(1))],
            "steam_appid": appid.group(1), "processes": [], "profile": "competitive",
        }
    return found


def all_games(config: dict) -> dict[str, dict]:
    return config["games"] | steam_games(config)


def requested_game(text: str, games: dict[str, dict], allow_guess: bool = False) -> str | None:
    words = normalize(text).split()
    compact = "".join(words)
    candidates: list[tuple[str, str]] = []
    for key, game in games.items():
        for alias in game.get("aliases", []):
            alias = normalize(alias)
            if alias in words or (len(alias) > 3 and alias.replace(" ", "") in compact):
                return key
            candidates.append((key, alias))
    if not allow_guess:
        return None
    winner, score = "", 0.0
    for word in words:
        for key, alias in candidates:
            ratio = SequenceMatcher(None, word, alias.replace(" ", "")).ratio()
            if ratio > score:
                winner, score = key, ratio
    if score >= 0.72:
        logging.info("Guessed game '%s' as '%s' (%.2f)", text, winner, score)
        return winner
    return None


def has_wake_word(text: str) -> bool:
    normalized = normalize(text)
    return any(re.search(rf"(?<!\w){re.escape(word)}(?!\w)", normalized) for word in WAKE_WORDS)


def is_game_request(text: str) -> bool:
    return bool(set(normalize(text).split()) & GAME_WORDS)


def requested_profile(text: str) -> str | None:
    words = set(normalize(text).split())
    return next((profile for profile, aliases in PROFILE_ALIASES.items() if words & aliases), None)


def is_status_request(text: str) -> bool:
    words = set(normalize(text).split())
    return bool(words & {"запущено", "играет", "какая", "какие", "статус"})


def is_close_request(text: str) -> bool:
    return bool(set(normalize(text).split()) & {"закрой", "закрыть", "выключи", "заверши", "останови", "выйти"})


def launch_game(game: dict) -> bool:
    try:
        if appid := game.get("steam_appid"):
            subprocess.Popen(["explorer.exe", f"steam://rungameid/{appid}"], close_fds=True)
            return True
        launcher = Path(game.get("launcher", ""))
        if not launcher.is_file():
            return False
        subprocess.Popen([str(launcher)], close_fds=True)
        return True
    except OSError as error:
        logging.warning("Could not start %s: %s", game_label(game), error)
        return False


def running_process_names() -> set[str]:
    """Read process names through Windows itself, with no third-party package."""
    output = subprocess.run(["tasklist", "/fo", "csv", "/nh"], capture_output=True,
                            text=True, encoding="utf-8", errors="ignore",
                            creationflags=subprocess.CREATE_NO_WINDOW).stdout
    return {row[0].lower() for row in csv.reader(output.splitlines()) if row}


def running_games(games: dict[str, dict]) -> list[tuple[str, dict]]:
    processes = running_process_names()
    found = []
    for key, game in games.items():
        expected = {name.lower() for name in game.get("processes", [])}
        if processes & expected:
            found.append((key, game))
    return found


def close_game(game: dict) -> bool:
    """Close only processes explicitly listed for a configured game."""
    expected = {name.lower() for name in game.get("processes", [])}
    targets = expected & running_process_names()
    if not targets:
        return False
    for process in targets:
        subprocess.run(["taskkill", "/im", process, "/t", "/f"], capture_output=True,
                       creationflags=subprocess.CREATE_NO_WINDOW)
    return True


def apply_profile(config: dict, profile_name: str) -> bool:
    profile = config["profiles"].get(profile_name)
    if not profile:
        return False
    try:
        subprocess.run(["powercfg", "/setactive", profile["power_scheme"]], check=False,
                       creationflags=subprocess.CREATE_NO_WINDOW)
        for program in profile.get("launch", []):
            executable = Path(program)
            if executable.is_file():
                subprocess.Popen([str(executable)], close_fds=True)
        logging.info("Applied profile: %s", profile_name)
        return True
    except OSError as error:
        logging.warning("Could not apply profile %s: %s", profile_name, error)
        return False


def discard_buffer(audio_queue: queue.Queue[bytes]) -> None:
    while True:
        try:
            audio_queue.get_nowait()
        except queue.Empty:
            return


def main() -> None:
    if not MODEL_DIR.is_dir():
        raise FileNotFoundError(f"Missing Vosk model: {MODEL_DIR}")
    config = load_config()
    SetLogLevel(-1)
    speaker = create_speaker()
    say(speaker, "Здравствуйте. Я Джарвис. Чем займемся, сэр?")

    audio_queue: queue.Queue[bytes] = queue.Queue()
    def microphone_callback(indata, frames, time_info, status) -> None:
        if status:
            logging.warning("Microphone status: %s", status)
        audio_queue.put(bytes(indata))

    sample_rate = int(sd.query_devices(kind="input")["default_samplerate"])
    recognizer = KaldiRecognizer(Model(str(MODEL_DIR)), sample_rate)
    cooldown_until = awaiting_game_choice_until = 0.0
    awaiting_command_until = time.monotonic() + 20
    active_game = ""
    last_auto_scan = 0.0

    with sd.RawInputStream(samplerate=sample_rate, blocksize=4_000, dtype="int16", channels=1, callback=microphone_callback):
        print("JARVIS is listening offline.")
        while True:
            data = audio_queue.get()
            now = time.monotonic()
            games = all_games(config)

            # Automatic game detection and profile activation, at most once a minute.
            if now - last_auto_scan >= 60:
                last_auto_scan = now
                detected = running_games(games)
                if detected and detected[0][0] != active_game:
                    active_game, detected_game = detected[0]
                    profile = detected_game.get("profile")
                    if profile:
                        apply_profile(config, profile)
                        logging.info("Auto profile %s for %s", profile, active_game)

            if not recognizer.AcceptWaveform(data):
                continue
            text = json.loads(recognizer.Result()).get("text", "")
            logging.info("Final recognition: %s", text)

            selected = requested_game(text, games, allow_guess=now < awaiting_game_choice_until)
            if now < awaiting_game_choice_until and selected:
                awaiting_game_choice_until = 0.0
                game = games[selected]
                if launch_game(game):
                    apply_profile(config, game.get("profile", "chill"))
                    say(speaker, "Открываю " + game_label(game) + ".")
                else:
                    say(speaker, "Я не нашел это приложение на компьютере.")
            elif now < awaiting_command_until:
                awaiting_command_until = 0.0
                profile = requested_profile(text)
                if is_status_request(text):
                    detected = running_games(games)
                    say(speaker, "Сейчас запущено: " + ", ".join(game_label(game) for _, game in detected) + "." if detected else "Сейчас игр не найдено.")
                elif is_close_request(text):
                    detected = running_games(games)
                    target = selected or (detected[0][0] if detected else None)
                    if target and close_game(games[target]):
                        say(speaker, "Закрываю " + game_label(games[target]) + ".")
                    else:
                        say(speaker, "Я не нашел запущенную игру для закрытия.")
                elif profile:
                    if apply_profile(config, profile):
                        say(speaker, "Включаю профиль " + config["profiles"][profile]["label"] + ".")
                    else:
                        say(speaker, "Не удалось применить профиль.")
                elif selected:
                    game = games[selected]
                    if launch_game(game):
                        apply_profile(config, game.get("profile", "chill"))
                        say(speaker, "Открываю " + game_label(game) + ".")
                    else:
                        say(speaker, "Я не нашел это приложение на компьютере.")
                elif is_game_request(text):
                    say(speaker, "Что открыть: Майнкрафт, Роблокс, Стим, Эпик или установленную игру из Стим?")
                    discard_buffer(audio_queue)
                    awaiting_game_choice_until = time.monotonic() + 20
                else:
                    say(speaker, "Пока я умею запускать и закрывать игры, проверять запущенную игру и включать игровые профили.")
            elif has_wake_word(text) and now >= cooldown_until:
                cooldown_until = now + 4
                say(speaker, "Чем займемся, сэр?")
                discard_buffer(audio_queue)
                awaiting_command_until = time.monotonic() + 20
            recognizer.Reset()


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        logging.exception("JARVIS could not start")
        print(f"JARVIS could not start: {error}", file=sys.stderr)
        raise
