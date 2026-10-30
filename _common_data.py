"""Local common-source connection; configured independently of business logic."""
from pathlib import Path
import json,sys
_base=Path(__file__).resolve().parent
while not (_base/'공통데이터_연결.json').is_file() and _base!=_base.parent:
    _base=_base.parent
_config=json.loads((_base/'공통데이터_연결.json').read_text(encoding='utf-8-sig'))
_code=str(Path(_config['root'])/'05_공통계산')
if _code not in sys.path:sys.path.insert(0,_code)
import common_sources
COMMON=common_sources.client(_config['project_id']) if _config.get('enabled') else None
