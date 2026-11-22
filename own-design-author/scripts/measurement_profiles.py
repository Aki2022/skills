"""The measurement profiles (own-design-author/references/measurement-profiles.json), read the same way by every Python
caller. A profile id is never spelled in code: an id the registry does not know is refused."""
import json
from pathlib import Path

REGISTRY = Path(__file__).resolve().parents[1] / 'references' / 'measurement-profiles.json'


def load():
    return json.loads(REGISTRY.read_text(encoding='utf-8'))


def profile_ids():
    return list(load()['profiles'])


def profile(name):
    profiles = load()['profiles']
    if not isinstance(name, str) or name not in profiles:
        raise ValueError('invalid rendering profile')
    return profiles[name]
