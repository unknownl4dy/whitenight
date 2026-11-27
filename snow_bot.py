"""
폭설 좀비 아포칼립스 디스코드 봇 (러너 vs 백귀)
설치: pip install discord.py openpyxl
실행: 환경변수 DISCORD_TOKEN 설정 후  python snow_bot.py
테스트용 엑셀 샘플 생성: python snow_bot.py sample   (snow.xlsx 가 만들어집니다)

환경변수 (Cloudtype에서 설정)
  DISCORD_TOKEN   : 봇 토큰 (필수)
  SHEET_ID        : 구글 시트 주소의 /d/ 와 /edit 사이 긴 문자열. 없으면 snow.xlsx 파일을 직접 읽고 씀
  SHEET_WEBHOOK   : Apps Script 웹 앱 주소 (시트에 체력·감염도 기록할 때 필요. SHEET_ID와 함께 설정)
  WEBHOOK_SECRET  : Apps Script 의 SECRET 과 똑같은 문자
  GM_CHANNEL_ID   : (선택) 총괄만 볼 수 있는 채널의 ID. 전투 후 감염 판정 결과가 여기에 올라옵니다

시트 탭 구조 (맨 윗줄은 제목줄)
  "러너"   : 이름 | 디스코드ID | 체력 | 최대체력(비우면 100) | 감염도(총괄 전용) | 조
  "백귀"   : 이름 | 체력 | 공격(주사위 면수, 비우면 20) | 조(비우면 모든 조)      ← 다음 전투에 나올 백귀를 적어두는 곳
  "아이템" : 아이템명 | 종류(회복/공격) | 값 | 소유자(디스코드ID 또는 러너 이름) | 남은횟수 | 비고
              · 모든 아이템은 1회용: 남은횟수를 비우거나 1로 적으면 한 번 쓰고 끝 (2 이상이면 그 횟수만큼)
  "자원"   : 맨 윗줄 약품 | 목재 | 고철, 둘째 줄에 공동 자원 개수 (모두가 함께 쓰는 자원, 총괄이 직접 관리)
  "웨이브" : 날짜(2026-11-24 형식) | 비율(예: 0.8) | 전체인원수(비우면 '러너' 탭 인원수)   ← 날마다 한 줄씩 총괄이 적음
  "방어벽" : 날짜 | 방어벽수 | 필요한백귀수 | 결과      ← 봇이 자동으로 만듭니다
  "운세기록" : 봇이 자동으로 만듭니다 (하루 횟수·아이템 지급 수 확인용)
  "감염기록" : 봇이 자동으로 만듭니다 (감염도가 오를 때마다 한 줄)

규칙 요약
  러너 체력 100 / 공격 D30(+공격 보정템) / 도주 D10 (7 이상 성공, 자기 차례에 가능)
  백귀 기본 체력 50 / 공격 D20 (시트·명령어로 바꿀 수 있음)
  전투: 러너 전원이 /공격 (또는 /도주) → 백귀가 한꺼번에 반격 → 다시 러너 턴 … 백귀가 모두 쓰러지면 자동 종료
  감염도(비공개, 전투가 끝난 뒤 판정): 풀데미지 피격 D20 / 체력 70↓ D10 / 50↓ D15 / 30↓ D20 (체력 구간은 가장 낮은 구간 하나만) / 조사마다 D5 (전날 방어벽 실패 시 D10)
"""
import asyncio
import io
import json
import os
import random
import re
import sys
import threading
import time
import urllib.request
import math
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from typing import Optional

import discord
from discord import app_commands
from openpyxl import Workbook, load_workbook

# ============================== 설정 ==============================
EXCEL_PATH = "snow.xlsx"
SHEET_ID = os.environ.get("SHEET_ID")
SHEET_WEBHOOK = os.environ.get("SHEET_WEBHOOK")
WEBHOOK_SECRET = os.environ.get("WEBHOOK_SECRET", "")
_gm = os.environ.get("GM_CHANNEL_ID", "").strip()
GM_CHANNEL_ID = int(_gm) if _gm.isdigit() else None

RUNNER_DICE = 30             # 러너 공격 다이스
DEFAULT_RUNNER_HP = 100      # 시트에 최대체력이 비어 있을 때
DEFAULT_ENEMY_HP = 50        # 백귀 기본 체력
DEFAULT_ENEMY_DICE = 20      # 백귀 기본 공격 다이스
FLEE_DIE, FLEE_NEED = 10, 7  # 도주: D10, 7 이상 성공
JUDGE_DIE, JUDGE_NEED = 10, 7  # /판정: D10, 7 이상 성공
MAX_RUNNERS = 5              # 한 전투에 참가할 수 있는 러너 수
PHASE_DELAY = 5              # 마지막 러너가 행동한 뒤 백귀가 반격하기까지 기다리는 시간(초)

# --- 감염도 (유저에게 비공개) ---
FULL_HIT_DIE = 20            # 풀데미지(백귀 공격 다이스의 최댓값)를 맞았을 때
FULL_HIT_ONCE = True         # True: 전투당 한 번만 굴림 / False: 풀데미지를 맞은 횟수만큼 굴림
HP_BRACKETS = [(30, 20), (50, 15), (70, 10)]   # (체력이 이 값 이하일 때, 굴릴 주사위 면수) - 가장 낮은 구간 하나만 적용
INVESTIGATE_DIE = 5          # 조사 한 번당 D5
INVESTIGATE_DIE_FAILED = 10  # 전날 방어벽 구축에 실패했다면 그날 조사는 D10

# --- 방어벽 ---
RES_ORDER = ["약품", "목재", "고철"]   # 방어벽은 반드시 이 순서로 반복해서 세움 (약품-목재-고철-약품-…)
BARRIER_MAX_PER_USE = 40     # 한 번에 만들 수 있는 최대 개수
BARRIER_WINDOW = None        # 방어벽 생성 시간을 제한하려면 ("23:20", "23:30") 처럼 적기. None 이면 항상 가능
SHOW_ZOMBIE_COUNT = False    # True 면 방어벽 결과 카드에 그날 들어오는 백귀 수도 보여줌 (기본은 성공/실패만)
KST = timezone(timedelta(hours=9))

# --- 운세 (라디오 운세) ---
FORTUNE_DAILY_LIMIT = 1         # 한 사람이 하루(한국 시간)에 볼 수 있는 횟수. 0이면 무제한
FORTUNE_CUT_CHANCE = 0.03       # 감염도가 줄어들 확률 (공개 문구는 '머리가 맑아진다' 같은 말뿐)
FORTUNE_CUT_RANGE = (1, 2)      # 줄어드는 감염도 양 (이 범위에서 무작위)
FORTUNE_ITEM_CHANCE = 0.05      # 아이템을 얻을 확률 (17명 기준 하루 평균 1개 안팎, 하루 총 2개까지)
FORTUNE_ITEM_DAILY_CAP = 2      # 하루에 운세로 나갈 수 있는 아이템 총 개수 (전체 합계, 먼저 뽑은 사람이 가져감)
# (이름, 종류, 값, 가중치) — 공격템의 값은 공격 보정, 회복템의 값은 회복량. 모두 1회용
FORTUNE_ITEMS = [
    ("부러진 쇠지렛대", "공격", 3, 35),
    ("날 선 손도끼", "공격", 5, 15),
    ("마른 건빵 한 봉지", "회복", 10, 35),
    ("온기가 남은 통조림", "회복", 20, 15),
]
FORTUNE_TITLE = "📻 라디오 운세"
FORTUNE_INTRO = "낡은 라디오의 다이얼을 천천히 돌린다. 지직거리는 잡음 사이로 누군가의 목소리가 끼어든다."
# (등급, 가중치, 한마디들) — 오늘의 운세 방송 느낌. 문구는 마음껏 추가/수정하세요
FORTUNE_RANKS = [
    ("대길 🌟", 8, ["오늘은 뭘 해도 술술 풀리는 날! 평소 미뤄둔 일에 도전해 보세요.", "별자리고 뭐고 오늘은 당신이 1위입니다. 눈보라도 비켜 갈 기세예요."]),
    ("길 ✨", 24, ["좋은 일이 슬쩍 다가오는 하루. 작은 행운을 놓치지 마세요.", "오늘은 운이 좋을 것 같습니다. 기분 좋게 시작해 보세요."]),
    ("중길 🙂", 28, ["무난하고 평온한 하루. 평범함이 가장 큰 행운일지도요.", "크게 좋지도 나쁘지도 않은 날. 꾸준함이 열쇠입니다."]),
    ("소흉 😐", 24, ["살짝 삐걱거리는 날. 서두르면 손해를 봅니다.", "오늘은 운이 조금 기울었어요. 한 번 더 확인하고 움직이세요."]),
    ("흉 ☁️", 14, ["오늘은 운이 나쁠 것 같습니다. 조용히 몸을 사리세요.", "먹구름 주의보! 괜한 일에 휘말리지 않도록 조심하세요."]),
    ("대흉 ⚡", 2, ["오늘 라디오를 켠 당신, 용기가 대단합니다. 부디 무사하시길…", "이럴 수가, 오늘은 가만히 있는 게 최고의 전략입니다."]),
]
FORTUNE_ADVICE = [
    "친구와 말다툼은 금물! 한마디 참으면 두 마디가 편해집니다.",
    "남의 말을 끝까지 들어주면 뜻밖의 도움을 받아요.",
    "오늘은 먼저 인사를 건네 보세요. 따뜻한 말 한마디가 돌아옵니다.",
    "배가 고프면 예민해져요. 끼니는 거르지 마세요.",
    "발밑을 조심하세요. 눈길은 늘 미끄럽습니다.",
    "혼자 끙끙대지 말고 동료에게 털어놓으면 길이 보입니다.",
    "오늘은 괜히 큰소리 치지 않기. 작은 약속이 큰 신뢰가 됩니다.",
    "잠깐의 휴식이 큰 실수를 막아 줍니다. 눈 좀 붙이세요.",
    "물건을 아끼세요. 오늘 아낀 한 개가 내일의 한 끼가 됩니다.",
    "뒤를 돌아보지 마세요. 아무것도 없을 거예요. …아마도요.",
    "오늘은 웃음이 최고의 방한복입니다. 많이 웃으세요.",
    "급할수록 돌아가세요. 지름길엔 눈이 더 쌓여 있어요.",
    "누군가 부탁하면 흔쾌히 도와주세요. 복이 되어 돌아옵니다.",
    "고민은 따뜻한 걸 마시면서 하면 절반으로 줄어듭니다.",
    "오늘은 동쪽에서 귀인이 나타납니다. 다만 동쪽이 어디인지는 모르겠네요.",
    "소지품 점검! 잃어버리기 쉬운 날입니다.",
    "너무 늦은 밤에는 돌아다니지 마세요. 어둠 속에서는 방향 감각이 흐려집니다.",
    "오늘의 연애운은 눈이 녹는 속도와 비슷합니다. 천천히, 하지만 확실하게.",
    "말보다 행동! 오늘은 한 걸음이 열 마디보다 낫습니다.",
    "뜻밖의 칭찬을 듣게 될지도 몰라요. 겸손하게 받아들이세요.",
]
FORTUNE_LUCKY = [
    "마른 양말", "따뜻한 차 한 잔", "녹슨 열쇠", "찢어진 지도", "빨간 털장갑", "반쯤 남은 초", "주머니 속 사탕",
    "접힌 종이학", "낡은 손목시계", "작은 나침반", "파란 목도리", "깨진 거울 조각", "성냥 한 개비", "동그란 단추",
    "오래된 엽서", "따끈한 수프", "고장 난 손전등", "털모자", "끈 풀린 군화", "차가운 금속 손잡이",
]
FORTUNE_DIRECTION = ["동쪽", "서쪽", "남쪽", "북쪽", "창가 쪽", "난로 옆", "계단 근처", "문 앞"]
FORTUNE_FOUND = "그런데 오늘은 운이 아주 좋았다! 길을 걷다 눈 속에서 쓸 만한 물건을 주웠다."
FORTUNE_CLEAR = [       # 감염도가 줄었을 때 (감염이라는 말은 쓰지 않음)
    "방송을 듣고 있자니 머리가 맑아지는 기분이 든다.",
    "익숙한 노래가 흘러나온다. 기분이 한결 좋아진다.",
    "라디오 너머의 목소리에 마음이 가벼워진다.",
]

ATTACK_VERBS = ["물어뜯는다", "할퀸다", "달려들어 할퀸다", "목덜미를 노리고 덤빈다", "으르렁거리며 물어뜯는다"]

COLOR_INTRO, COLOR_ATK, COLOR_ENEMY, COLOR_END, COLOR_INFO = 0x5DADE2, 0x2ECC71, 0xC0392B, 0xF1C40F, 0x95A5A6


def d(n):
    return random.randint(1, n)


# ============================== 데이터 ==============================
class StoreError(Exception):
    """시트 읽기/쓰기 실패 (사용자에게 보여줄 문구가 들어 있음)"""


@dataclass
class Runner:
    name: str
    uid: int
    hp: int
    max_hp: int
    group: Optional[int] = None
    fled: bool = False
    acted: bool = False
    full_hits: int = 0           # 풀데미지를 맞은 횟수 (감염 판정용)

    @property
    def alive(self): return self.hp > 0
    @property
    def active(self): return self.hp > 0 and not self.fled


@dataclass
class Enemy:
    name: str
    hp: int
    max_hp: int
    dice: int = DEFAULT_ENEMY_DICE

    @property
    def alive(self): return self.hp > 0


@dataclass
class Battle:
    channel: object
    group: Optional[int]
    runners: dict
    enemies: list
    first: str                   # "러너"(조우) / "백귀"(발각)
    round: int = 0
    phase: str = "runner"        # runner / enemy / done
    ended: bool = False
    lock: asyncio.Lock = field(default_factory=asyncio.Lock)


battles: dict = {}               # 채널ID -> Battle


# ============================== 표시 도우미 ==============================
def hp_bar(cur, mx, width=10):
    mx = max(1, mx)
    filled = round(width * max(0, cur) / mx)
    if cur > 0 and filled == 0:
        filled = 1
    return "▰" * filled + "▱" * (width - filled)


def hp_line(cur, mx):
    return f"{hp_bar(cur, mx)} **{max(0, cur)}/{mx}**"


def clip(text, limit=1000):
    return text if len(text) <= limit else text[:limit - 1] + "…"


def runners_text(b):
    lines = []
    for r in b.runners.values():
        tag = " 🏃도주" if r.fled else (" 💀쓰러짐" if not r.alive else "")
        lines.append(f"**{r.name}** {hp_line(r.hp, r.max_hp)}{tag}")
    return "\n".join(lines) or "-"


def enemies_text(b):
    lines = []
    for n, e in enumerate(b.enemies, 1):
        tag = "" if e.alive else " 💀"
        lines.append(f"`{n}` **{e.name}** {hp_line(e.hp, e.max_hp)} · 공격 D{e.dice}{tag}")
    return "\n".join(lines) or "-"


def pending_names(b):
    return [r.name for r in b.runners.values() if r.active and not r.acted]


def find_enemy(b, text):
    """이름 / 번호 / 자동완성 글자('1. 남성 백귀 (50/50)') / 띄어쓰기 달라도 / 이름 앞부분 으로 백귀를 찾음"""
    t = (text or "").strip()
    t = re.sub(r"\s*\(\s*\d+\s*/\s*\d+\s*\)\s*$", "", t)          # 끝의 (50/50) 제거
    m = re.match(r"^(\d+)\s*[.)]\s*(.*)$", t)                           # 앞의 '1.' 제거
    if m:
        t = m.group(2).strip() or m.group(1)
    if not t:
        return None
    for e in b.enemies:
        if e.name == t:
            return e
    if t.isdigit() and 1 <= int(t) <= len(b.enemies):
        return b.enemies[int(t) - 1]
    flat = re.sub(r"\s+", "", t)
    for e in b.enemies:
        if re.sub(r"\s+", "", e.name) == flat:
            return e
    cands = [e for e in b.enemies if e.alive and flat in re.sub(r"\s+", "", e.name)]
    return cands[0] if len(cands) == 1 else None


# ============================== 시트 연결 ==============================
def _txt(v):
    return "" if v is None else str(v).strip()


def _int(v, default=0):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return default


def _uid(v):
    s = _txt(v)
    if not s:
        return None
    try:
        return int(float(s)) if "e" in s.lower() or "." in s else int(s)
    except ValueError:
        return None


def remaining(kind, cell_value):
    """남은 횟수. 모든 아이템은 비우면 1회용 (2 이상을 적으면 그 횟수만큼 사용 가능)"""
    s = _txt(cell_value)
    if s == "":
        return 1
    return _int(s, 0)


def owner_match(cell_value, uid, name):
    s = _txt(cell_value)
    return s != "" and (s == str(uid) or s == name)


def barrier_plan(start, n, have):
    """방어벽을 start개 세운 상태에서 n개를 더 세울 때: (필요 자원, 지금 자원으로 만들 수 있는 최대 개수)"""
    need = {r: 0 for r in RES_ORDER}
    for k in range(start, start + n):
        need[RES_ORDER[k % 3]] += 1
    can, used = 0, {r: 0 for r in RES_ORDER}
    for k in range(start, start + n):
        r = RES_ORDER[k % 3]
        if used[r] + 1 > have.get(r, 0):
            break
        used[r] += 1; can += 1
    return need, can


def barrier_shortage_text(start, n, have):
    need, can = barrier_plan(start, n, have)
    lack = [f"{r} {need[r] - have.get(r, 0)}개 부족(필요 {need[r]}, 보유 {have.get(r, 0)})" for r in RES_ORDER if need[r] > have.get(r, 0)]
    return f"자원이 부족합니다: {', '.join(lack)}. 지금 공동 자원으로는 최대 {can}개까지 만들 수 있습니다. (다음 순서: {RES_ORDER[start % 3]})"


def zombies_needed(people, ratio):
    """그날 들어오는 백귀 수 = 전체 인원수 × 비율 (소수점은 올림)"""
    return math.ceil(people * ratio - 1e-9)


def date_text(v):
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    return _txt(v)


def find_tab(wb, name):
    """탭 이름 앞뒤에 공백이 붙어 있어도 찾아줌"""
    for n in wb.sheetnames:
        if n.strip() == name:
            return wb[n]
    return None


def parse_resources(wb):
    """공동 자원 {약품, 목재, 고철}: '자원' 탭 첫 줄이 제목, 둘째 줄이 개수"""
    ws = find_tab(wb, "자원")
    if ws is None:
        return None
    rows = list(ws.iter_rows(min_row=1, max_row=2, values_only=True))
    head = [_txt(x) for x in rows[0]] if rows else []
    vals = rows[1] if len(rows) > 1 else ()
    return {r: _int(vals[head.index(r)] if r in head and head.index(r) < len(vals) else 0, 0) for r in RES_ORDER}


def parse_wave(wb, date):
    """'웨이브' 탭에서 오늘 날짜 줄 → (비율, 인원수 직접지정 또는 None). 못 찾으면 StoreError(이유를 자세히 알려줌)"""
    ws = find_tab(wb, "웨이브")
    if ws is None:
        raise StoreError(f"시트에 '웨이브' 탭이 없습니다. (지금 시트에 있는 탭: {', '.join(wb.sheetnames)})")
    seen = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or _txt(row[0]) == "":
            continue
        dt = date_text(row[0])
        seen.append(dt)
        if dt != date:
            continue
        try:
            ratio = float(str(row[1]).strip().replace(",", ".")) if len(row) > 1 and _txt(row[1]) != "" else None
        except ValueError:
            ratio = None
        if ratio is None:
            raise StoreError(f"'웨이브' 탭의 {date} 줄은 찾았는데 B열 '비율'이 비어 있거나 숫자가 아닙니다.")
        ppl = _int(row[2], 0) if len(row) > 2 and _txt(row[2]) != "" else None
        return ratio, (ppl or None)
    raise StoreError(f"'웨이브' 탭에서 오늘({date}) 날짜 줄을 못 찾았습니다. 봇이 읽은 날짜: "
                     + (", ".join(seen[:8]) if seen else "(없음)") + " — A열 날짜가 맞는지 확인해 주세요.")


def barrier_status(wb, date):
    """오늘의 방어 상황 계산 (읽기 전용): 인원, 비율, 필요한 백귀 수, 현재 방어벽 수"""
    ratio, ppl = parse_wave(wb, date)
    people = ppl or len(parse_runners(wb))
    built = 0
    if "방어벽" in wb.sheetnames:
        for row in wb["방어벽"].iter_rows(min_row=2, values_only=True):
            if row and date_text(row[0]) == date:
                built = _int(row[1] if len(row) > 1 else 0, 0)
    return {"date": date, "people": people, "ratio": ratio, "required": zombies_needed(people, ratio), "built": built}


def prev_barrier_failed(wb, today):
    """전날(한국 시간) 방어벽이 실패였는지: 전날 웨이브 줄이 있고 방어벽 수가 필요한 백귀 수보다 적으면 True"""
    y = (datetime.strptime(today, "%Y-%m-%d") - timedelta(days=1)).strftime("%Y-%m-%d")
    try:
        st = barrier_status(wb, y)
    except StoreError:
        return False
    return st["built"] < st["required"]


def parse_runners(wb):
    out = []
    for row in wb["러너"].iter_rows(min_row=2, values_only=True):
        if not row or not _txt(row[0]):
            continue
        uid = _uid(row[1] if len(row) > 1 else None)
        mx = _int(row[3] if len(row) > 3 else None, DEFAULT_RUNNER_HP) or DEFAULT_RUNNER_HP
        hp = _int(row[2] if len(row) > 2 else None, mx) if _txt(row[2] if len(row) > 2 else None) != "" else mx
        grp = _int(row[5], 0) if len(row) > 5 and _txt(row[5]) != "" else None
        out.append({"name": _txt(row[0]), "uid": uid, "hp": hp, "max_hp": mx, "group": grp})
    return out


def parse_enemies(wb):
    out = []
    if "백귀" not in wb.sheetnames:
        return out
    for row in wb["백귀"].iter_rows(min_row=2, values_only=True):
        if not row or not _txt(row[0]):
            continue
        hp = _int(row[1], DEFAULT_ENEMY_HP) if len(row) > 1 and _txt(row[1]) != "" else DEFAULT_ENEMY_HP
        dice = _int(row[2], DEFAULT_ENEMY_DICE) if len(row) > 2 and _txt(row[2]) != "" else DEFAULT_ENEMY_DICE
        grp = _int(row[3], 0) if len(row) > 3 and _txt(row[3]) != "" else None
        out.append({"name": _txt(row[0]), "hp": hp, "dice": dice or DEFAULT_ENEMY_DICE, "group": grp})
    return out


def parse_items(wb):
    out = []
    if "아이템" not in wb.sheetnames:
        return out
    for row in wb["아이템"].iter_rows(min_row=2, values_only=True):
        if not row or not _txt(row[0]):
            continue
        kind = _txt(row[1]) if len(row) > 1 else ""
        out.append({"name": _txt(row[0]), "kind": kind, "value": _int(row[2] if len(row) > 2 else 0),
                    "owner": _txt(row[3]) if len(row) > 3 else "",
                    "left": remaining(kind, row[4] if len(row) > 4 else None)})
    return out


class SheetStore:
    """구글 시트: 읽기는 xlsx 내보내기, 쓰기는 Apps Script 웹 앱 (체력·감염도·아이템을 시트에서 바로 계산)"""

    def _wb(self):
        url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"
        with urllib.request.urlopen(url, timeout=20) as resp:
            return load_workbook(io.BytesIO(resp.read()), data_only=True)

    def _post(self, payload):
        payload["secret"] = WEBHOOK_SECRET
        req = urllib.request.Request(SHEET_WEBHOOK, data=json.dumps(payload).encode("utf-8"),
                                     headers={"Content-Type": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                res = json.loads(resp.read().decode("utf-8"))
        except Exception as e:
            raise StoreError(f"시트에 연결하지 못했습니다 ({e})")
        if not res.get("ok"):
            raise StoreError(res.get("error", "알 수 없는 오류"))
        return res

    def load_runners(self): return parse_runners(self._wb())
    def load_enemies(self): return parse_enemies(self._wb())
    def load_items(self): return parse_items(self._wb())
    def load_lookup(self):
        wb = self._wb(); return parse_runners(wb), parse_items(wb)
    def set_hp(self, uid, hp): return self._post({"action": "hp", "uid": str(uid), "hp": hp})
    def hp_delta(self, uid, delta): return self._post({"action": "hp_delta", "uid": str(uid), "delta": delta})
    def transfer(self, owner_uid, target_uid, item):
        return self._post({"action": "transfer", "owner": str(owner_uid), "uid": str(target_uid), "item": item})
    def fortune(self, uid, date, limit, cap, cut, item):
        return self._post({"action": "fortune", "uid": str(uid), "date": date, "limit": limit, "cap": cap, "cut": cut,
                           "item": ({"name": item[0], "kind": item[1], "value": item[2]} if item else None)})
    def load_resources(self): return parse_resources(self._wb())
    def barrier(self, n, date): return self._post({"action": "barrier", "n": n, "date": date})
    def barrier_remove(self, date, count, refund):
        return self._post({"action": "barrier_remove", "date": date, "count": count or 0, "refund": bool(refund)})
    def barrier_status(self, date): return barrier_status(self._wb(), date)
    def prev_failed(self, today): return prev_barrier_failed(self._wb(), today)
    def infect(self, uid, add, note): return self._post({"action": "infect", "uid": str(uid), "add": add, "note": note})
    def heal(self, target_uid, owner_uid, item):
        return self._post({"action": "heal", "uid": str(target_uid), "owner": str(owner_uid), "item": item})
    def use_item(self, owner_uid, item):
        return self._post({"action": "use_item", "owner": str(owner_uid), "item": item})


class LocalStore:
    """snow.xlsx 파일을 직접 읽고 씀 (SHEET_ID 가 없을 때: 테스트·연습용)"""
    lock = threading.Lock()

    def _wb(self, data_only=True):
        return load_workbook(EXCEL_PATH, data_only=data_only)

    def load_runners(self): return parse_runners(self._wb())
    def load_enemies(self): return parse_enemies(self._wb())
    def load_items(self): return parse_items(self._wb())
    def load_lookup(self):
        wb = self._wb(); return parse_runners(wb), parse_items(wb)

    @staticmethod
    def _runner_row(ws, uid):
        for r in range(2, ws.max_row + 1):
            if _uid(ws.cell(r, 2).value) == uid and _txt(ws.cell(r, 1).value):
                return r
        raise StoreError("명단에서 러너를 찾을 수 없습니다")

    def set_hp(self, uid, hp):
        with self.lock:
            wb = self._wb(False); ws = wb["러너"]; r = self._runner_row(ws, uid)
            mx = _int(ws.cell(r, 4).value, DEFAULT_RUNNER_HP) or DEFAULT_RUNNER_HP
            ws.cell(r, 3).value = max(0, min(mx, hp)); wb.save(EXCEL_PATH)
            return {"ok": True}

    def load_resources(self): return parse_resources(self._wb())
    def barrier_status(self, date): return barrier_status(self._wb(), date)
    def prev_failed(self, today): return prev_barrier_failed(self._wb(), today)

    def fortune(self, uid, date, limit, cap, cut, item):
        with self.lock:
            wb = self._wb(False)
            rs = wb["러너"]; r = self._runner_row(rs, uid)
            name = _txt(rs.cell(r, 1).value)
            if "운세기록" not in wb.sheetnames:
                wb.create_sheet("운세기록").append(["시간", "날짜", "이름", "디스코드ID", "감염감소", "아이템"])
            fs = wb["운세기록"]
            used, items_today = 0, 0
            for x in range(2, fs.max_row + 1):
                if date_text(fs.cell(x, 2).value) == date:
                    if _txt(fs.cell(x, 4).value) == str(uid):
                        used += 1
                    if _txt(fs.cell(x, 6).value):
                        items_today += 1
            if limit and used >= limit:
                raise StoreError("오늘의 라디오 운세는 이미 들었습니다. 내일 다시 해 보세요.")
            given = None
            if item and items_today < cap:
                if "아이템" not in wb.sheetnames:
                    raise StoreError("시트에 '아이템' 탭이 없습니다")
                wb["아이템"].append([item[0], item[1], item[2], name, 1, f"운세 {date}"])
                given = item[0]
            cut_done = 0
            if cut > 0:
                cur = _int(rs.cell(r, 5).value, 0)
                cut_done = min(cur, cut)
                if cut_done > 0:
                    rs.cell(r, 5).value = cur - cut_done
                    if "감염기록" not in wb.sheetnames:
                        wb.create_sheet("감염기록").append(["시간", "이름", "디스코드ID", "증가", "합계", "사유"])
                    wb["감염기록"].append([time.strftime("%Y-%m-%d %H:%M:%S"), name, str(uid), -cut_done, cur - cut_done, "운세"])
            fs.append([time.strftime("%Y-%m-%d %H:%M:%S"), date, name, str(uid), cut_done, given or ""])
            wb.save(EXCEL_PATH)
            return {"ok": True, "count": used + 1, "item_given": given is not None, "cut": cut_done}

    def transfer(self, owner_uid, target_uid, item):
        with self.lock:
            wb = self._wb(False)
            runners = {x["uid"]: x for x in parse_runners(wb)}
            if owner_uid not in runners:
                raise StoreError("명단에서 러너를 찾을 수 없습니다")
            if target_uid not in runners:
                raise StoreError("받는 사람이 러너 명단에 없습니다")
            found = self._find_item(wb, owner_uid, runners[owner_uid]["name"], item)
            if not found:
                raise StoreError(f"'{item}' 을(를) 갖고 있지 않거나 이미 다 썼습니다")
            ws, r, k, left = found
            ws.cell(r, 4).value = runners[target_uid]["name"]
            wb.save(EXCEL_PATH)
            return {"ok": True, "kind": k, "value": _int(ws.cell(r, 3).value), "left": left, "to": runners[target_uid]["name"]}

    def barrier_remove(self, date, count, refund):
        with self.lock:
            wb = self._wb(False)
            bs = find_tab(wb, "방어벽")
            row = next((x for x in range(2, bs.max_row + 1) if date_text(bs.cell(x, 1).value) == date), None) if bs else None
            built = _int(bs.cell(row, 2).value, 0) if row else 0
            if built <= 0:
                raise StoreError("오늘 만든 방어벽이 없습니다")
            k = built if not count or count >= built else count
            new = built - k
            refunded = {r: 0 for r in RES_ORDER}
            if refund:
                for idx in range(new, built):
                    refunded[RES_ORDER[idx % 3]] += 1
                ws = find_tab(wb, "자원")
                if ws is not None:
                    head = [_txt(c.value) for c in ws[1]]
                    for r in RES_ORDER:
                        col = head.index(r) + 1
                        ws.cell(2, col).value = _int(ws.cell(2, col).value, 0) + refunded[r]
            required = _int(bs.cell(row, 3).value, 0)
            bs.cell(row, 2).value = new
            bs.cell(row, 4).value = "성공" if new >= required else "실패"
            wb.save(EXCEL_PATH)
            return {"ok": True, "removed": k, "total": new, "required": required, "refunded": refunded if refund else None,
                    "success": new >= required}

    def barrier(self, n, date):
        with self.lock:
            wb = self._wb(False)
            st = barrier_status(wb, date)
            have = parse_resources(wb)
            if have is None:
                raise StoreError("시트에 '자원' 탭이 없습니다")
            start = st["built"]
            need, can = barrier_plan(start, n, have)
            if can < n:
                raise StoreError(barrier_shortage_text(start, n, have))
            ws = find_tab(wb, "자원")
            head = [_txt(c.value) for c in ws[1]]
            for r in RES_ORDER:
                col = head.index(r) + 1
                ws.cell(2, col).value = have[r] - need[r]
            total = start + n
            success = total >= st["required"]
            if "방어벽" not in wb.sheetnames:
                wb.create_sheet("방어벽").append(["날짜", "방어벽수", "필요한백귀수", "결과"])
            bs = wb["방어벽"]
            row = next((x for x in range(2, bs.max_row + 1) if date_text(bs.cell(x, 1).value) == date), None)
            vals = [date, total, st["required"], "성공" if success else "실패"]
            if row:
                for c, v in enumerate(vals, 1):
                    bs.cell(row, c).value = v
            else:
                bs.append(vals)
            wb.save(EXCEL_PATH)
            return {"ok": True, "total": total, "required": st["required"], "success": success, "used": need,
                    "left": {r: have[r] - need[r] for r in RES_ORDER}, "next": RES_ORDER[total % 3]}

    def hp_delta(self, uid, delta):
        with self.lock:
            wb = self._wb(False); ws = wb["러너"]; r = self._runner_row(ws, uid)
            mx = _int(ws.cell(r, 4).value, DEFAULT_RUNNER_HP) or DEFAULT_RUNNER_HP
            hp = max(0, min(mx, _int(ws.cell(r, 3).value, mx) + delta))
            ws.cell(r, 3).value = hp; wb.save(EXCEL_PATH)
            return {"ok": True, "hp": hp, "max": mx, "name": ws.cell(r, 1).value}

    def infect(self, uid, add, note):
        with self.lock:
            wb = self._wb(False); ws = wb["러너"]; r = self._runner_row(ws, uid)
            total = _int(ws.cell(r, 5).value, 0) + add
            ws.cell(r, 5).value = total
            if "감염기록" not in wb.sheetnames:
                wb.create_sheet("감염기록").append(["시간", "이름", "디스코드ID", "증가", "합계", "사유"])
            wb["감염기록"].append([time.strftime("%Y-%m-%d %H:%M:%S"), ws.cell(r, 1).value, str(uid), add, total, note])
            wb.save(EXCEL_PATH)
            return {"ok": True, "total": total}

    @staticmethod
    def _find_item(wb, owner_uid, owner_name, item, kind=None):
        ws = wb["아이템"]
        for r in range(2, ws.max_row + 1):
            if _txt(ws.cell(r, 1).value) != item:
                continue
            k = _txt(ws.cell(r, 2).value)
            if kind and k != kind:
                continue
            if not owner_match(ws.cell(r, 4).value, owner_uid, owner_name):
                continue
            left = remaining(k, ws.cell(r, 5).value)
            if left is None or left > 0:
                return ws, r, k, left
        return None

    def _consume(self, ws, r, k, left):
        if left is not None:
            ws.cell(r, 5).value = left - 1
            return left - 1
        return None

    def use_item(self, owner_uid, item):
        with self.lock:
            wb = self._wb(False)
            name = next((x["name"] for x in parse_runners(wb) if x["uid"] == owner_uid), "")
            found = self._find_item(wb, owner_uid, name, item, "공격")
            if not found:
                raise StoreError(f"'{item}' 공격 보정 아이템을 갖고 있지 않거나 이미 다 썼습니다")
            ws, r, k, left = found
            val = _int(ws.cell(r, 3).value)
            new_left = self._consume(ws, r, k, left)
            wb.save(EXCEL_PATH)
            return {"ok": True, "kind": k, "value": val, "left": new_left}

    def heal(self, target_uid, owner_uid, item):
        with self.lock:
            wb = self._wb(False)
            runners = {x["uid"]: x for x in parse_runners(wb)}
            if owner_uid not in runners:
                raise StoreError("명단에서 러너를 찾을 수 없습니다")
            found = self._find_item(wb, owner_uid, runners[owner_uid]["name"], item, "회복")
            if not found:
                raise StoreError(f"'{item}' 회복템을 갖고 있지 않거나 이미 다 썼습니다")
            ws_i, ri, k, left = found
            ws = wb["러너"]; r = self._runner_row(ws, target_uid)
            cur = _int(ws.cell(r, 3).value, 0)
            mx = _int(ws.cell(r, 4).value, DEFAULT_RUNNER_HP) or DEFAULT_RUNNER_HP
            if cur <= 0:
                raise StoreError("체력이 0인 러너는 회복할 수 없습니다")
            val = _int(ws_i.cell(ri, 3).value)
            new_hp = min(mx, cur + val)
            ws.cell(r, 3).value = new_hp
            new_left = self._consume(ws_i, ri, k, left)
            wb.save(EXCEL_PATH)
            return {"ok": True, "hp": new_hp, "max": mx, "healed": new_hp - cur, "value": val,
                    "left": new_left, "name": ws.cell(r, 1).value}


store = SheetStore() if (SHEET_ID and SHEET_WEBHOOK) else LocalStore()
if SHEET_ID and not SHEET_WEBHOOK:
    print("⚠ SHEET_ID 는 있는데 SHEET_WEBHOOK 이 없어 시트 쓰기가 안 됩니다. snow.xlsx 파일 모드로 동작합니다.")

_cache = {"items": [], "runners": [], "started": False}


async def cache_loop():
    """자동완성은 3초 안에 답해야 해서, 시트(러너·아이템)를 20초마다 미리 읽어 둡니다"""
    while True:
        try:
            runners, items = await asyncio.to_thread(store.load_lookup)
            _cache["runners"], _cache["items"] = runners, items
        except Exception as e:
            print("자동완성용 시트 읽기 실패:", e)
        await asyncio.sleep(20)


# ============================== 감염도 판정 ==============================
def judge_infection(r: Runner):
    """전투가 끝난 뒤 한 번 호출. (증가량, 설명 목록) 반환"""
    parts, total = [], 0
    if r.full_hits:
        for _ in range(1 if FULL_HIT_ONCE else r.full_hits):
            v = d(FULL_HIT_DIE); total += v; parts.append(f"풀데미지 피격 D{FULL_HIT_DIE} → {v}")
    for limit, die in HP_BRACKETS:                       # 가장 낮은 구간 하나만 (HP_BRACKETS 는 낮은 값부터)
        if r.hp <= limit:
            v = d(die); total += v; parts.append(f"체력 {r.hp} (≤{limit}) D{die} → {v}")
            break
    return total, parts


# ============================== 전투 진행 ==============================
async def write_hp(b, runners):
    """바뀐 체력을 시트에 바로 반영. 실패하면 채널에 알림 (총괄이 시트를 직접 고치도록)"""
    results = await asyncio.gather(*[asyncio.to_thread(store.set_hp, r.uid, r.hp) for r in runners],
                                   return_exceptions=True)
    bad = [r.name for r, res in zip(runners, results) if isinstance(res, Exception)]
    if bad:
        await b.channel.send(f"⚠ 시트에 체력을 반영하지 못했습니다: {', '.join(bad)} (총괄님이 시트를 직접 확인해 주세요)")


async def end_battle(b: Battle, outcome, judge=True):
    if b.ended:
        return
    b.ended, b.phase = True, "done"
    battles.pop(b.channel.id, None)
    colors = {"승리": COLOR_END}
    em = discord.Embed(title=f"전투 종료 — {outcome}", color=colors.get(outcome, COLOR_INFO),
                       description={"승리": "백귀가 모두 쓰러졌다.", "전멸": "러너가 모두 쓰러졌다.",
                                    "도주": "러너가 모두 전장을 벗어났다."}.get(outcome, "총괄이 전투를 종료했다."))
    em.add_field(name="러너", value=clip(runners_text(b)), inline=False)
    em.add_field(name="백귀", value=clip(enemies_text(b)), inline=False)
    await b.channel.send(embed=em)
    if judge:
        await settle_infection(b)


async def settle_infection(b: Battle):
    """전투 후 감염 판정: 시트에 바로 더하고, 결과는 총괄 채널에만 알림 (러너에게는 공개하지 않음)"""
    lines = []
    for r in b.runners.values():
        add, parts = judge_infection(r)
        if add <= 0:
            lines.append(f"**{r.name}** (체력 {r.hp}): 증가 없음")
            continue
        try:
            res = await asyncio.to_thread(store.infect, r.uid, add, "전투 후 판정: " + " / ".join(parts))
            lines.append(f"**{r.name}** +{add}  →  감염도 {res.get('total', '?')}\n　" + " / ".join(parts))
        except Exception as e:
            lines.append(f"**{r.name}** +{add} ⚠ 시트 반영 실패({e}) — 직접 더해 주세요\n　" + " / ".join(parts))
    if GM_CHANNEL_ID:
        ch = client.get_channel(GM_CHANNEL_ID)
        if ch:
            em = discord.Embed(title="🧪 감염 판정 (총괄 전용)", description=clip("\n".join(lines), 4000), color=0x8E44AD)
            em.set_footer(text=f"{getattr(b.channel, 'name', '')} 전투 · 시트 '러너' 탭의 감염도에 이미 반영됨")
            await ch.send(embed=em)


def check_runner_side_over(b: Battle):
    """러너가 더 싸울 수 없으면 종료 사유를 반환"""
    if any(r.active for r in b.runners.values()):
        return None
    return "도주" if any(r.alive for r in b.runners.values()) else "전멸"


async def enemy_phase(b: Battle):
    b.phase = "enemy"
    if PHASE_DELAY:
        await asyncio.sleep(PHASE_DELAY)
    lines, hit = [], []
    for e in b.enemies:
        if not e.alive:
            continue
        targets = [r for r in b.runners.values() if r.active]
        if not targets:
            break
        t = random.choice(targets)
        roll = d(e.dice)
        crit = e.dice > 1 and roll == e.dice
        before = t.hp
        t.hp = max(0, t.hp - roll)
        if crit:
            t.full_hits += 1
        if t not in hit:
            hit.append(t)
        lines.append(f"👹 **{e.name}** → **{t.name}** {random.choice(ATTACK_VERBS)}!\n"
                     f"　🎲 D{e.dice} → **{roll}**" + (" 💥 **치명타!**" if crit else "")
                     + f"  ·  체력 {before} → **{t.hp}**" + ("  💀 쓰러졌다" if not t.alive else ""))
    em = discord.Embed(title=f"[{b.round}라운드] 백귀의 반격", description=clip("\n".join(lines) or "백귀가 움직이지 않았다.", 4000),
                       color=COLOR_ENEMY)
    em.add_field(name="러너 체력", value=clip(runners_text(b)), inline=False)
    await b.channel.send(embed=em)
    if hit:
        await write_hp(b, hit)
    over = check_runner_side_over(b)
    if over:
        await end_battle(b, over)


async def begin_round(b: Battle):
    b.round += 1
    for r in b.runners.values():
        r.acted = False
    if b.first == "백귀":
        await enemy_phase(b)
        if b.ended:
            return
    b.phase = "runner"
    if b.round == 1 and b.first == "러너":
        return                                          # 첫 라운드는 시작 안내 카드가 대신함
    em = discord.Embed(title=f"[{b.round}라운드] 러너 턴", color=COLOR_INTRO,
                       description="`/공격 대상:백귀이름` 또는 `/도주` 로 행동하세요. (공격 대상은 반드시 지정)")
    em.add_field(name="백귀", value=clip(enemies_text(b)), inline=False)
    em.add_field(name="행동할 러너", value=", ".join(pending_names(b)) or "-", inline=False)
    await b.channel.send(embed=em)


async def runners_done(b: Battle):
    """러너 전원이 행동을 마쳤을 때: (러너 선공이면) 백귀 반격 → 정산 → 다음 라운드"""
    b.phase = "enemy"                                   # 처리 중에는 새 행동을 받지 않음
    if b.first == "러너":
        await enemy_phase(b)
        if b.ended:
            return
    em = discord.Embed(title=f"[{b.round}라운드] 정산", color=COLOR_INFO)
    em.add_field(name="러너", value=clip(runners_text(b)), inline=False)
    em.add_field(name="백귀", value=clip(enemies_text(b)), inline=False)
    await b.channel.send(embed=em)
    await begin_round(b)


async def after_action(b: Battle):
    """공격/도주 직후 호출 (b.lock 을 잡은 상태)"""
    if all(not e.alive for e in b.enemies):
        await end_battle(b, "승리")
        return
    over = check_runner_side_over(b)
    if over:
        await end_battle(b, over)
        return
    if not pending_names(b):
        await runners_done(b)


# ============================== 명령어 처리 ==============================
async def private(i, text):
    await i.followup.send(text, ephemeral=True)


async def do_attack(i, target_text, note, item):
    await i.response.defer(ephemeral=True)
    b = battles.get(i.channel_id)
    if not b:
        return await private(i, "이 채널에는 진행 중인 전투가 없습니다.")
    async with b.lock:
        r = b.runners.get(i.user.id)
        if b.ended or b.phase != "runner":
            return await private(i, "지금은 러너 턴이 아닙니다.")
        if not r:
            return await private(i, "이 전투의 참가자가 아닙니다.")
        if not r.alive:
            return await private(i, "쓰러진 상태라 행동할 수 없습니다.")
        if r.fled:
            return await private(i, "이미 전장을 벗어났습니다.")
        if r.acted:
            return await private(i, "이번 라운드의 행동을 이미 마쳤습니다.")
        e = find_enemy(b, target_text)
        if not e or not e.alive:
            alive = ", ".join(f"{x.name} (번호 {n})" for n, x in enumerate(b.enemies, 1) if x.alive)
            return await private(i, f"공격 대상을 찾을 수 없습니다. `대상` 칸에 백귀 이름이나 번호를 입력하세요.\n지금 대상: {alive}")
        bonus, item_txt = 0, ""
        if item:
            try:
                res = await asyncio.to_thread(store.use_item, i.user.id, item)
            except StoreError as ex:
                return await private(i, str(ex))
            if res.get("kind") != "공격":
                return await private(i, f"'{item}' 은(는) 공격 보정 아이템이 아닙니다.")
            bonus = res.get("value", 0)
            item_txt = f" +{bonus}({item})"
        roll = d(RUNNER_DICE)
        total = roll + bonus
        before = e.hp
        e.hp = max(0, e.hp - total)
        r.acted = True
        em = discord.Embed(title=f"⚔️ {r.name}의 공격", color=COLOR_ATK,
                           description=(f"*{note}*\n\n" if note else "") + f"**{e.name}** 을(를) 공격했다!")
        dice_txt = f"D{RUNNER_DICE} → **{roll}**" + (f"{item_txt} = **{total}**" if item_txt else "")
        em.add_field(name="공격 다이스", value=dice_txt, inline=True)
        em.add_field(name=f"{e.name} 체력", value=hp_line(e.hp, e.max_hp) + ("  💀 쓰러졌다" if not e.alive else ""), inline=False)
        left = pending_names(b)
        if left:
            em.set_footer(text="남은 행동자: " + ", ".join(left))
        await b.channel.send(embed=em)
        await private(i, "✅ 공격했습니다.")
        await after_action(b)


async def do_flee(i):
    await i.response.defer(ephemeral=True)
    b = battles.get(i.channel_id)
    if not b:
        return await private(i, "이 채널에는 진행 중인 전투가 없습니다.")
    async with b.lock:
        r = b.runners.get(i.user.id)
        if b.ended or b.phase != "runner":
            return await private(i, "지금은 러너 턴이 아닙니다.")
        if not r:
            return await private(i, "이 전투의 참가자가 아닙니다.")
        if not r.active:
            return await private(i, "쓰러졌거나 이미 전장을 벗어났습니다.")
        if r.acted:
            return await private(i, "이번 라운드의 행동을 이미 마쳤습니다.")
        roll = d(FLEE_DIE)
        ok = roll >= FLEE_NEED
        r.acted = True
        if ok:
            r.fled = True
        em = discord.Embed(title=f"🏃 {r.name}의 도주 시도", color=COLOR_INFO,
                           description=f"D{FLEE_DIE} → **{roll}** (≥{FLEE_NEED} 성공)\n"
                           + ("**도주 성공!** 전장을 벗어났다." if ok else "**도주 실패…** 그대로 공격을 받게 된다."))
        left = pending_names(b)
        if left:
            em.set_footer(text="남은 행동자: " + ", ".join(left))
        await b.channel.send(embed=em)
        await private(i, "✅ 도주를 시도했습니다.")
        await after_action(b)


def uid_in_battle(uid):
    return any(uid in b.runners and b.runners[uid].alive and not b.runners[uid].fled for b in battles.values())


async def do_heal(i, item, target_member):
    await i.response.defer(ephemeral=True)
    target = target_member or i.user
    if uid_in_battle(i.user.id) or uid_in_battle(target.id):
        return await private(i, "전투 중에는 식량으로 회복할 수 없습니다. 정산이 끝난 뒤 사용하세요.")
    try:
        res = await asyncio.to_thread(store.heal, target.id, i.user.id, item)
    except StoreError as ex:
        return await private(i, str(ex))
    who = res.get("name") or target.display_name
    em = discord.Embed(title="🍖 체력 회복", color=COLOR_ATK,
                       description=f"**{i.user.display_name}** 이(가) **{item}** 을(를) "
                                   + ("먹었다." if target.id == i.user.id else f"**{who}** 에게 건넸다."))
    em.add_field(name="체력", value=f"{hp_line(res['hp'], res['max'])}  (+{res['healed']})", inline=False)
    await i.channel.send(embed=em)
    left = res.get("left")
    await private(i, "✅ 사용했습니다." + ("" if left is None else f" (남은 횟수 {left})"))


async def do_transfer(i, item, target):
    await i.response.defer(ephemeral=True)
    if getattr(target, "bot", False) or target.id == i.user.id:
        return await private(i, "다른 러너를 지정하세요.")
    if uid_in_battle(i.user.id) or uid_in_battle(target.id):
        return await private(i, "전투 중에는 아이템을 양도할 수 없습니다. 전투가 끝난 뒤 사용하세요.")
    try:
        res = await asyncio.to_thread(store.transfer, i.user.id, target.id, item)
    except StoreError as ex:
        return await private(i, str(ex))
    me = next((x for x in _cache["runners"] if x["uid"] == i.user.id), None)
    giver = me["name"] if me else i.user.display_name
    icon = "🍖" if res.get("kind") == "회복" else "🗡️"
    em = discord.Embed(title="🎁 아이템 양도", color=COLOR_INFO,
                       description=f"**{giver}** 이(가) {icon} **{item}** 을(를) **{res.get('to') or target.display_name}** 에게 건넸다.")
    await i.channel.send(embed=em)
    await private(i, "✅ 양도했습니다.")
    try:                                                   # 자동완성 목록이 바로 바뀌도록 새로 읽어 둠
        _cache["runners"], _cache["items"] = await asyncio.to_thread(store.load_lookup)
    except Exception:
        pass


async def do_fortune(i):
    await i.response.defer(ephemeral=True)
    today = datetime.now(KST).strftime("%Y-%m-%d")
    cut = random.randint(*FORTUNE_CUT_RANGE) if random.random() < FORTUNE_CUT_CHANCE else 0
    item = None
    if random.random() < FORTUNE_ITEM_CHANCE:
        item = random.choices(FORTUNE_ITEMS, weights=[x[3] for x in FORTUNE_ITEMS])[0]
    try:
        res = await asyncio.to_thread(store.fortune, i.user.id, today, FORTUNE_DAILY_LIMIT, FORTUNE_ITEM_DAILY_CAP, cut, item)
    except StoreError as ex:
        return await private(i, str(ex))
    me = next((x for x in _cache["runners"] if x["uid"] == i.user.id), None)
    name = me["name"] if me else i.user.display_name
    lucky = bool(res.get("item_given") or cut)
    pool = FORTUNE_RANKS[:2] if lucky else FORTUNE_RANKS
    rank = random.choices(pool, weights=[x[1] for x in pool])[0]
    lines = [FORTUNE_INTRO, "", f"📻 **오늘의 운세: {rank[0]}**", random.choice(rank[2]),
             f"💬 {random.choice(FORTUNE_ADVICE)}",
             f"🍀 행운의 물건: **{random.choice(FORTUNE_LUCKY)}** · 행운의 방향: **{random.choice(FORTUNE_DIRECTION)}**"]
    color = COLOR_END if res.get("item_given") else COLOR_INFO
    if res.get("item_given"):
        icon = "🗡️" if item[1] == "공격" else "🍖"
        lines.append(f"\n✨ {FORTUNE_FOUND}\n{icon} **{item[0]}** ({'공격 +' if item[1] == '공격' else '회복 +'}{item[2]}, 1회용) — `/내정보`에서 확인하세요.")
    if cut:                                                # 감염이라는 말은 숨기고 분위기 문구만
        lines.append("\n" + random.choice(FORTUNE_CLEAR))
    em = discord.Embed(title=f"{FORTUNE_TITLE} — {name}", description="\n".join(lines), color=color)
    await i.channel.send(embed=em)
    await private(i, "✅ 라디오를 켰습니다.")
    if res.get("item_given"):
        try:                                               # 아이템 자동완성 목록이 바로 갱신되도록
            _cache["runners"], _cache["items"] = await asyncio.to_thread(store.load_lookup)
        except Exception:
            pass


async def do_barrier(i, n):
    await i.response.defer(ephemeral=True)
    if n < 1 or n > BARRIER_MAX_PER_USE:
        return await private(i, f"방어벽 수는 1~{BARRIER_MAX_PER_USE} 사이로 입력하세요.")
    if BARRIER_WINDOW:
        now = datetime.now(KST).strftime("%H:%M")
        if not (BARRIER_WINDOW[0] <= now < BARRIER_WINDOW[1]):
            return await private(i, f"방어벽은 {BARRIER_WINDOW[0]} ~ {BARRIER_WINDOW[1]} 사이에만 만들 수 있습니다.")
    today = datetime.now(KST).strftime("%Y-%m-%d")
    me = next((x for x in _cache["runners"] if x["uid"] == i.user.id), None)
    name = me["name"] if me else i.user.display_name
    try:
        res = await asyncio.to_thread(store.barrier, n, today)
    except StoreError as ex:
        if str(ex).startswith("자원이 부족합니다"):               # 자원 부족은 모두가 볼 수 있게 채널에 공개
            em = discord.Embed(title="🛡️ 방어벽 설치 실패", color=COLOR_ENEMY,
                               description=f"**{name}** 이(가) 방어벽 **{n}개** 를 세우려 했지만 자원이 모자랐다.\n\n{ex}")
            await i.channel.send(embed=em)
            return await private(i, "⚠️ 자원이 부족해 만들지 못했습니다.")
        return await private(i, str(ex))
    ok = res["success"]
    em = discord.Embed(title="🛡️ 방어벽 설치", color=COLOR_ATK if ok else COLOR_ENEMY,
                       description=f"**{name}** 이(가) 방어벽 **{n}개** 를 세웠다.")
    em.add_field(name="방어 결과",
                 value=("✅ **방어 성공** — 방어벽이 몰려오는 백귀를 모두 막아냈다." if ok
                        else "⚠️ **방어 실패** — 막지 못한 백귀가 거점으로 넘어왔다.\n[명일 조사 시 감염 및 백귀 조우 확률이 증가합니다.]")
                 + (f"\n오늘 들어오는 백귀 {res['required']}마리" if SHOW_ZOMBIE_COUNT else ""), inline=False)
    em.add_field(name="사용한 자원", value=" · ".join(f"{k} {v}" for k, v in res["used"].items() if v) or "-", inline=True)
    em.add_field(name="오늘 방어벽", value=f"총 **{res['total']}개**\n다음 순서: {res['next']}", inline=True)
    em.add_field(name="남은 공동 자원", value=" · ".join(f"{k} {v}" for k, v in res["left"].items()), inline=False)
    em.set_footer(text="방어벽은 약품 → 목재 → 고철 순서로 쌓이며, 하루가 지나면 사라집니다")
    await i.channel.send(embed=em)
    await private(i, "✅ 방어벽을 만들었습니다.")


async def do_barrier_remove(i, count, refund):
    await i.response.defer(ephemeral=True)
    today = datetime.now(KST).strftime("%Y-%m-%d")
    if count is not None and count < 1:
        return await private(i, "삭제할 개수는 1 이상으로 입력하세요. (비우면 오늘 방어벽을 전부 삭제)")
    try:
        res = await asyncio.to_thread(store.barrier_remove, today, count, refund)
    except StoreError as ex:
        return await private(i, str(ex))
    text = f"✅ 오늘({today}) 방어벽 {res['removed']}개를 삭제했습니다. 현재 {res['total']}개."
    if res.get("refunded") is not None:
        text += "\n자원 반환: " + (" · ".join(f"{k} {v}" for k, v in res["refunded"].items() if v) or "없음")
    else:
        text += "\n(자원은 돌려주지 않았습니다. 필요하면 시트 '자원' 탭을 직접 고치세요.)"
    await private(i, text)


async def do_barrier_status(i):
    await i.response.defer(ephemeral=True)
    today = datetime.now(KST).strftime("%Y-%m-%d")
    try:
        st = await asyncio.to_thread(store.barrier_status, today)
        have = await asyncio.to_thread(store.load_resources)
    except StoreError as ex:
        return await private(i, str(ex))
    except Exception as ex:
        return await private(i, f"시트를 읽지 못했습니다: {ex}")
    short = st["required"] - st["built"]
    em = discord.Embed(title=f"🛡️ 오늘({today}) 방어 현황 (총괄 전용)", color=COLOR_INFO)
    em.add_field(name="들어오는 백귀", value=f"**{st['required']}마리**\n(인원 {st['people']} × 비율 {st['ratio']:g})", inline=True)
    em.add_field(name="현재 방어벽", value=f"**{st['built']}개**\n" + ("✅ 충분" if short <= 0 else f"⚠️ {short}개 부족"), inline=True)
    if have is not None:
        em.add_field(name="공동 자원", value=" · ".join(f"{k} {v}" for k, v in have.items()), inline=False)
    em.set_footer(text=f"다음에 쌓을 자원: {RES_ORDER[st['built'] % 3]}")
    await i.followup.send(embed=em, ephemeral=True)


async def do_judge(i, note):
    await i.response.defer(ephemeral=True)
    roll = d(JUDGE_DIE)
    ok = roll >= JUDGE_NEED
    name = i.user.display_name
    me = next((x for x in _cache["runners"] if x["uid"] == i.user.id), None)
    if me:
        name = me["name"]
    em = discord.Embed(title=f"🎲 {name}의 판정", color=COLOR_ATK if ok else COLOR_ENEMY,
                       description=(f"*{note}*\n\n" if note else "") + ("**✅ 성공**" if ok else "**❌ 실패**"))
    em.add_field(name="판정 다이스", value=f"D{JUDGE_DIE} → **{roll}**  (≥{JUDGE_NEED} 성공)", inline=False)
    await i.channel.send(embed=em)
    await private(i, "✅ 판정했습니다.")


async def do_info(i):
    await i.response.defer(ephemeral=True)
    try:
        runners = await asyncio.to_thread(store.load_runners)
        items = await asyncio.to_thread(store.load_items)
    except Exception as e:
        return await private(i, f"시트를 읽지 못했습니다: {e}")
    me = next((x for x in runners if x["uid"] == i.user.id), None)
    if not me:
        return await private(i, "명단에 없는 사용자입니다. 총괄에게 문의하세요.")
    mine = [x for x in items if owner_match(x["owner"], i.user.id, me["name"]) and (x["left"] is None or x["left"] > 0)]
    em = discord.Embed(title=f"🧍 {me['name']}", color=COLOR_INFO)
    em.add_field(name="체력", value=hp_line(me["hp"], me["max_hp"]), inline=False)
    try:
        rs = await asyncio.to_thread(store.load_resources)
    except Exception:
        rs = None
    if rs is not None:
        em.add_field(name="공동 자원", value=" · ".join(f"{k} {v}" for k, v in rs.items()), inline=False)
    if mine:
        em.add_field(name="아이템", value=clip("\n".join(
            f"{'🍖' if x['kind'] == '회복' else '🗡️'} **{x['name']}** "
            f"({'회복 +' if x['kind'] == '회복' else '공격 +'}{x['value']}"
            + ("" if x["left"] is None else f", 남은 {x['left']}회") + ")" for x in mine)), inline=False)
    else:
        em.add_field(name="아이템", value="없음", inline=False)
    await i.followup.send(embed=em, ephemeral=True)


async def investigate_die():
    """어제 방어벽이 실패했으면 D10, 아니면 D5 (시트를 못 읽으면 기본값)"""
    today = datetime.now(KST).strftime("%Y-%m-%d")
    try:
        return INVESTIGATE_DIE_FAILED if await asyncio.to_thread(store.prev_failed, today) else INVESTIGATE_DIE
    except Exception:
        return INVESTIGATE_DIE


async def do_investigate(i):
    await i.response.defer(ephemeral=True)
    die = await investigate_die()
    roll = d(die)
    try:
        await asyncio.to_thread(store.infect, i.user.id, roll, f"조사 시작 D{die} → {roll}")
    except StoreError as ex:
        return await private(i, str(ex))
    await i.channel.send(f"🔍 **{i.user.display_name}** 이(가) 조사를 시작한다.")
    await private(i, "✅ 조사를 시작했습니다.")


async def do_investigate_group(i, group):
    await i.response.defer(ephemeral=True)
    runners = [x for x in await asyncio.to_thread(store.load_runners) if x["group"] == group and x["uid"]]
    if not runners:
        return await private(i, f"{group}조 러너가 없습니다.")
    done, bad = [], []
    die = await investigate_die()
    for x in runners:
        roll = d(die)
        try:
            await asyncio.to_thread(store.infect, x["uid"], roll, f"조사 시작 D{die} → {roll}")
            done.append(x["name"])
        except StoreError as ex:
            bad.append(f"{x['name']}({ex})")
    await i.channel.send(f"🔍 **{group}조** 조사 시작 — " + ", ".join(done))
    await private(i, "✅ 처리했습니다." + (f" 실패: {', '.join(bad)}" if bad else ""))


async def do_hp_adjust(i, delta, member, group, reason):
    """총괄: 조사 중 피해/회복 등을 시트 체력에 바로 더하거나 뺌 (전투 중인 러너는 전투 체력도 같이 바뀜)"""
    await i.response.defer(ephemeral=True)
    if (member is None) == (group is None):
        return await private(i, "`대상자` 와 `조` 중 하나만 입력하세요.")
    if delta == 0:
        return await private(i, "증감 값이 0 입니다.")
    if member is not None:
        targets = [(member.id, None)]
    else:
        rows = await asyncio.to_thread(store.load_runners)
        targets = [(x["uid"], x["name"]) for x in rows if x["group"] == group and x["uid"]]
        if not targets:
            return await private(i, f"{group}조 러너가 없습니다.")
    lines, bad = [], []
    for uid, nm in targets:
        b = next((x for x in battles.values() if uid in x.runners), None)
        try:
            if b:                                      # 전투 중: 전투 체력과 시트를 같이 맞춤
                async with b.lock:
                    r = b.runners[uid]
                    r.hp = max(0, min(r.max_hp, r.hp + delta))
                    await asyncio.to_thread(store.set_hp, uid, r.hp)
                    res = {"hp": r.hp, "max": r.max_hp, "name": r.name}
            else:
                res = await asyncio.to_thread(store.hp_delta, uid, delta)
        except StoreError as ex:
            bad.append(f"{nm or uid}({ex})"); continue
        lines.append(f"**{res.get('name') or nm}**  {hp_line(res['hp'], res['max'])}" + ("  💀 쓰러졌다" if res["hp"] <= 0 else ""))
    if lines:
        sign = f"+{delta}" if delta > 0 else str(delta)
        em = discord.Embed(title=("🍖 체력 회복" if delta > 0 else "🩸 체력 감소") + f" ({sign})",
                           description=(f"*{reason}*\n\n" if reason else "") + "\n".join(lines),
                           color=COLOR_ATK if delta > 0 else COLOR_ENEMY)
        await i.channel.send(embed=em)
    await private(i, "✅ 반영했습니다." + (f" 실패: {', '.join(bad)}" if bad else ""))


async def do_start(i, group, first, enemy_count, hp, dice, picked):
    await i.response.defer(ephemeral=True)
    if i.channel_id in battles and not battles[i.channel_id].ended:
        return await private(i, "이 채널에서 이미 전투가 진행 중입니다.")
    try:
        rows = await asyncio.to_thread(store.load_runners)
        erows = await asyncio.to_thread(store.load_enemies)
    except Exception as e:
        return await private(i, f"시트를 읽지 못했습니다: {e}")
    notes = []
    if picked:
        chosen = []
        for m in picked:
            x = next((x for x in rows if x["uid"] == m.id), None)
            if not x:
                return await private(i, f"{m.display_name}님은 러너 명단에 없습니다.")
            chosen.append(x)
    else:
        chosen = [x for x in rows if x["group"] == group and x["uid"]]
    if not chosen:
        return await private(i, f"{group}조 러너가 없습니다. 시트의 '조' 열을 확인하세요.")
    busy = [x["name"] for x in chosen if uid_in_battle(x["uid"])]
    if busy:
        return await private(i, f"이미 다른 전투 중인 러너가 있습니다: {', '.join(busy)}")
    downed = [x["name"] for x in chosen if x["hp"] <= 0]
    chosen = [x for x in chosen if x["hp"] > 0]
    if downed:
        notes.append(f"체력 0이라 제외: {', '.join(downed)}")
    if not chosen:
        return await private(i, "참가할 수 있는 러너가 없습니다 (모두 체력 0).")
    if len(chosen) > MAX_RUNNERS:
        return await private(i, f"러너가 {len(chosen)}명입니다. 전투는 최대 {MAX_RUNNERS}명까지라 `러너1`~`러너{MAX_RUNNERS}` 로 직접 지정해 주세요.")
    if enemy_count and enemy_count > 0:
        # 시트에 적힌 백귀 중에서 무작위로 뽑음: 이 조로 지정된 백귀 먼저, 모자라면 조가 빈 백귀, 그래도 모자라면 기본 백귀
        mine = [x for x in erows if group and x["group"] == group]
        free = [x for x in erows if x["group"] is None] if group else [x for x in erows]   # 조 없이 시작하면 시트 전체에서 무작위
        random.shuffle(mine); random.shuffle(free)
        pool = (mine + free)[:enemy_count]
        enemies = []
        for x in pool:
            h = hp or x["hp"]
            enemies.append(Enemy(x["name"], h, h, dice or x["dice"]))
        for n in range(len(enemies) + 1, enemy_count + 1):
            enemies.append(Enemy(f"백귀{n}", hp or DEFAULT_ENEMY_HP, hp or DEFAULT_ENEMY_HP, dice or DEFAULT_ENEMY_DICE))
        names = [e.name for e in enemies]                      # 이름이 겹치면 번호를 붙여 구분
        seen = {}
        for e in enemies:
            if names.count(e.name) > 1:
                seen[e.name] = seen.get(e.name, 0) + 1
                e.name = f"{e.name}{seen[e.name]}"
    else:
        src = [x for x in erows if x["group"] in (None, group)] if group else erows
        if not src:
            return await private(i, "시트 '백귀' 탭에 이번 조에 해당하는 줄이 없습니다. `백귀수` 를 입력하거나 시트에 적어 주세요.")
        names = [x["name"] for x in src]
        enemies = []
        seen = {}
        for x in src:
            nm = x["name"]
            if names.count(nm) > 1:
                seen[nm] = seen.get(nm, 0) + 1
                nm = f"{nm}{seen[nm]}"
            enemies.append(Enemy(nm, x["hp"], x["hp"], x["dice"]))
    runners = {x["uid"]: Runner(x["name"], x["uid"], x["hp"], x["max_hp"], x["group"]) for x in chosen}
    b = Battle(i.channel, group, runners, enemies, first)
    battles[i.channel_id] = b
    async with b.lock:
        em = discord.Embed(title="❄️ 백귀와의 전투", color=COLOR_INTRO,
                           description=("**조우!** 러너가 먼저 움직인다." if first == "러너" else "**발각!** 백귀가 먼저 덮쳐 온다.")
                           + "\n\n`/공격 대상:백귀이름 지문:(선택)` · `/도주`\n공격 대상을 반드시 지정해 주세요. 러너 전원이 행동하면 백귀가 반격합니다.")
        em.add_field(name="러너", value=clip(runners_text(b)), inline=False)
        em.add_field(name="백귀", value=clip(enemies_text(b)), inline=False)
        em.add_field(name="행동 순서", value="러너 공격 → 백귀 반격 → 정산" if first == "러너" else "백귀 반격 → 러너 공격 → 정산", inline=False)
        await i.channel.send(embed=em)
        await private(i, "✅ 전투를 시작했습니다." + ("\n" + "\n".join(notes) if notes else ""))
        await begin_round(b)


async def do_force(i):
    await i.response.defer(ephemeral=True)
    b = battles.get(i.channel_id)
    if not b:
        return await private(i, "이 채널에는 진행 중인 전투가 없습니다.")
    async with b.lock:
        if b.ended or b.phase != "runner":
            return await private(i, "지금은 러너 턴이 아닙니다.")
        skipped = pending_names(b)
        for r in b.runners.values():
            r.acted = True
        if skipped:
            await b.channel.send(f"⏩ 총괄이 진행시켰습니다. 행동하지 않은 러너: {', '.join(skipped)}")
        await private(i, "✅ 진행했습니다.")
        await after_action(b)


async def do_stop(i, judge):
    await i.response.defer(ephemeral=True)
    b = battles.get(i.channel_id)
    if not b:
        return await private(i, "이 채널에는 진행 중인 전투가 없습니다.")
    async with b.lock:
        await end_battle(b, "총괄 종료", judge=judge)
        await private(i, "✅ 전투를 종료했습니다." + (" (감염 판정 포함)" if judge else " (감염 판정 없음)"))


async def do_enemy_edit(i, target_text, hp, dice):
    await i.response.defer(ephemeral=True)
    b = battles.get(i.channel_id)
    if not b:
        return await private(i, "이 채널에는 진행 중인 전투가 없습니다.")
    async with b.lock:
        e = find_enemy(b, target_text)
        if not e:
            return await private(i, "그 이름의 백귀가 없습니다.")
        if hp is not None:
            e.hp = max(0, hp); e.max_hp = max(e.max_hp, e.hp)
        if dice is not None and dice > 0:
            e.dice = dice
        await b.channel.send(f"🛠️ 백귀 **{e.name}** 의 상태가 바뀌었다.\n{enemies_text(b)}")
        await private(i, "✅ 수정했습니다.")
        if all(not x.alive for x in b.enemies):
            await end_battle(b, "승리")


# ============================== 디스코드 연결 ==============================
intents = discord.Intents.default()
client = discord.Client(intents=intents)
tree = app_commands.CommandTree(client)


@client.event
async def on_ready():
    if not _cache["started"]:
        _cache["started"] = True
        asyncio.create_task(cache_loop())
    await tree.sync()
    print("ready")


async def enemy_choices(i, current):
    b = battles.get(i.channel_id)
    if not b:
        return []
    cur = (current or "").strip()
    out = []
    for n, e in enumerate(b.enemies, 1):
        if e.alive and (not cur or cur in e.name or cur == str(n)):
            out.append(app_commands.Choice(name=f"{n}. {e.name} ({e.hp}/{e.max_hp})", value=e.name))
    return out[:25]


async def item_choices(i, current, kind):
    cur = (current or "").strip()
    me = next((x for x in _cache["runners"] if x["uid"] == i.user.id), None)
    names = {n for n in (me["name"] if me else None, i.user.display_name, getattr(i.user, "name", None),
                         getattr(i.user, "global_name", None)) if n}
    out, seen = [], set()
    for x in _cache["items"]:
        if (kind and x["kind"] != kind) or x["name"] in seen:
            continue
        if not (_txt(x["owner"]) == str(i.user.id) or _txt(x["owner"]) in names):   # 소유자 칸: 내 ID, 시트의 러너 이름, 디스코드 닉네임 중 하나
            continue
        if x["left"] is not None and x["left"] <= 0:
            continue
        if cur and cur not in x["name"]:
            continue
        seen.add(x["name"])
        out.append(app_commands.Choice(name=x["name"], value=x["name"]))
    return out[:25]


@tree.command(name="공격", description="백귀를 공격합니다 (D30)")
@app_commands.describe(대상="공격할 백귀 (이름 또는 번호)", 지문="행동 묘사 (선택)", 아이템="사용할 공격 보정 아이템 (선택, 내 것만)")
async def cmd_attack(i: discord.Interaction, 대상: str, 지문: str = None, 아이템: str = None):
    await do_attack(i, 대상, 지문, 아이템)


@cmd_attack.autocomplete("대상")
async def ac_target(i: discord.Interaction, current: str):
    return await enemy_choices(i, current)


@cmd_attack.autocomplete("아이템")
async def ac_atk_item(i: discord.Interaction, current: str):
    return await item_choices(i, current, "공격")


@tree.command(name="도주", description="전투에서 도주를 시도합니다 (D10, 7 이상 성공)")
async def cmd_flee(i: discord.Interaction):
    await do_flee(i)


@tree.command(name="회복", description="회복 아이템을 사용합니다 (전투 중 불가)")
@app_commands.describe(아이템="사용할 회복 아이템 이름 (내 것만)", 대상자="다른 러너에게 쓸 때만 지정 (비우면 나)")
async def cmd_heal(i: discord.Interaction, 아이템: str, 대상자: discord.Member = None):
    await do_heal(i, 아이템, 대상자)


@cmd_heal.autocomplete("아이템")
async def ac_heal_item(i: discord.Interaction, current: str):
    return await item_choices(i, current, "회복")


@tree.command(name="방어벽", description="공동 자원으로 방어벽을 만듭니다 (약품→목재→고철 순서로 자동 차감)")
@app_commands.describe(방어벽수="만들 방어벽 개수 (대표 한 명이 입력)")
async def cmd_barrier(i: discord.Interaction, 방어벽수: int):
    await do_barrier(i, 방어벽수)


@tree.command(name="방어벽삭제", description="[총괄] 오늘 만든 방어벽을 삭제합니다 (테스트·실수 정리용)")
@app_commands.default_permissions(manage_guild=True)
@app_commands.describe(개수="최근에 쌓은 것부터 이만큼만 삭제 (비우면 오늘 방어벽 전부)",
                       자원돌려주기="True 면 삭제한 방어벽에 쓰인 자원을 공동 자원에 돌려줌 (기본 False)")
async def cmd_barrier_remove(i: discord.Interaction, 개수: int = None, 자원돌려주기: bool = False):
    await do_barrier_remove(i, 개수, 자원돌려주기)


@tree.command(name="방어벽현황", description="[총괄] 오늘 들어오는 백귀 수와 방어벽 수를 확인합니다")
@app_commands.default_permissions(manage_guild=True)
async def cmd_barrier_status(i: discord.Interaction):
    await do_barrier_status(i)


@tree.command(name="운세", description="낡은 라디오로 오늘의 운세를 듣습니다 (하루 1회)")
async def cmd_fortune(i: discord.Interaction):
    await do_fortune(i)


@tree.command(name="판정", description="D10 판정을 합니다 (7 이상 성공)")
@app_commands.describe(지문="무엇을 판정하는지 적어 주세요 (선택)")
async def cmd_judge(i: discord.Interaction, 지문: str = None):
    await do_judge(i, 지문)


@tree.command(name="양도", description="내 아이템을 다른 러너에게 넘깁니다 (소유자가 바뀜)")
@app_commands.describe(아이템="넘길 아이템 이름 (내 것만)", 대상자="받을 러너")
async def cmd_transfer(i: discord.Interaction, 아이템: str, 대상자: discord.Member):
    await do_transfer(i, 아이템, 대상자)


@cmd_transfer.autocomplete("아이템")
async def ac_transfer_item(i: discord.Interaction, current: str):
    return await item_choices(i, current, None)


@tree.command(name="내정보", description="내 체력과 아이템을 확인합니다 (나에게만 보임)")
async def cmd_info(i: discord.Interaction):
    await do_info(i)


@tree.command(name="조사", description="조사를 시작합니다")
async def cmd_investigate(i: discord.Interaction):
    await do_investigate(i)


# ----- 총괄 전용 -----
@tree.command(name="자동전투", description="[총괄] 이 채널에서 백귀와의 자동 전투를 시작합니다")
@app_commands.default_permissions(manage_guild=True)
@app_commands.describe(조="참가할 조 번호 (시트의 '조' 열)", 선공="조우=러너 선공 / 발각=백귀 선공",
                       백귀수="비우면 시트 '백귀' 탭을 사용, 숫자를 넣으면 기본 백귀를 그 수만큼 만듦",
                       체력="(백귀수 사용 시) 백귀 체력, 기본 50", 공격="(백귀수 사용 시) 백귀 공격 주사위 면수, 기본 20",
                       러너1="직접 지정할 때 (조 대신)", 러너2="직접 지정", 러너3="직접 지정", 러너4="직접 지정", 러너5="직접 지정")
@app_commands.choices(선공=[app_commands.Choice(name="러너 선공 (조우)", value="러너"),
                            app_commands.Choice(name="백귀 선공 (발각)", value="백귀")])
async def cmd_start(i: discord.Interaction, 조: int, 선공: str = "러너", 백귀수: int = 0, 체력: int = None, 공격: int = None,
                    러너1: discord.Member = None, 러너2: discord.Member = None,
                    러너3: discord.Member = None, 러너4: discord.Member = None,
                    러너5: discord.Member = None):
    await do_start(i, 조, 선공, 백귀수, 체력, 공격, [m for m in (러너1, 러너2, 러너3, 러너4, 러너5) if m])


@tree.command(name="체력조정", description="[총괄] 러너의 체력을 더하거나 뺍니다 (조사 중 피해 등). 시트에 바로 반영")
@app_commands.default_permissions(manage_guild=True)
@app_commands.describe(증감="줄이려면 -10, 늘리려면 10 처럼 입력", 대상자="한 명에게 (조 대신)", 조="그 조 전체에게 (대상자 대신)", 사유="표시할 지문 (선택)")
async def cmd_hp_adjust(i: discord.Interaction, 증감: int, 대상자: discord.Member = None, 조: int = None, 사유: str = None):
    await do_hp_adjust(i, 증감, 대상자, 조, 사유)


@tree.command(name="강제진행", description="[총괄] 아직 행동하지 않은 러너를 건너뛰고 백귀 차례로 넘깁니다")
@app_commands.default_permissions(manage_guild=True)
async def cmd_force(i: discord.Interaction):
    await do_force(i)


@tree.command(name="전투종료", description="[총괄] 진행 중인 전투를 종료합니다")
@app_commands.default_permissions(manage_guild=True)
@app_commands.describe(감염판정="True 면 지금 체력 기준으로 감염 판정까지 함 (기본), False 면 판정 없이 종료")
async def cmd_stop(i: discord.Interaction, 감염판정: bool = True):
    await do_stop(i, 감염판정)


@tree.command(name="백귀수정", description="[총괄] 진행 중인 전투의 백귀 체력·공격을 바꿉니다")
@app_commands.default_permissions(manage_guild=True)
@app_commands.describe(대상="바꿀 백귀", 체력="새 체력 (현재 체력으로 설정)", 공격="새 공격 주사위 면수")
async def cmd_enemy_edit(i: discord.Interaction, 대상: str, 체력: int = None, 공격: int = None):
    await do_enemy_edit(i, 대상, 체력, 공격)


@cmd_enemy_edit.autocomplete("대상")
async def ac_edit_target(i: discord.Interaction, current: str):
    return await enemy_choices(i, current)


@tree.command(name="조사시작", description="[총괄] 한 조 전체의 조사를 시작합니다 (러너마다 감염도 D5 자동 상승)")
@app_commands.default_permissions(manage_guild=True)
async def cmd_investigate_group(i: discord.Interaction, 조: int):
    await do_investigate_group(i, 조)


# ============================== 샘플 시트 / 실행 ==============================
def make_sample():
    wb = Workbook()
    ws = wb.active; ws.title = "러너"
    ws.append(["이름", "디스코드ID", "체력", "최대체력", "감염도(총괄 전용)", "조"])
    ws.append(["러너A", "111111111111111111", 100, 100, 0, 1])      # ID는 반드시 텍스트로
    ws.append(["러너B", "222222222222222222", 100, 100, 0, 1])
    ws.append(["러너C", "333333333333333333", 100, 100, 0, 2])
    ws = wb.create_sheet("백귀")
    ws.append(["이름", "체력", "공격(주사위 면수)", "조(비우면 모든 조)"])
    ws.append(["백귀1", 50, 20, 1])
    ws.append(["백귀2", 50, 20, 1])
    ws = wb.create_sheet("자원")
    ws.append(["약품", "목재", "고철"])
    ws.append([6, 5, 5])
    ws = wb.create_sheet("웨이브")
    ws.append(["날짜", "비율", "전체인원수(비우면 러너 탭 인원수)"])
    ws.append([datetime.now(KST).strftime("%Y-%m-%d"), 1.2, 14])
    ws = wb.create_sheet("아이템")
    ws.append(["아이템명", "종류(회복/공격)", "값", "소유자(ID 또는 이름)", "남은횟수", "비고"])
    ws.append(["통조림", "회복", 20, "111111111111111111", 1, "1회용"])
    ws.append(["녹슨 쇠파이프", "공격", 5, "러너B", 1, "1회용"])
    for col in ("B",):
        for row in wb["러너"][col]:
            row.number_format = "@"
    wb.save(EXCEL_PATH)
    print(f"{EXCEL_PATH} 샘플 생성 완료")


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "sample":
        make_sample()
    else:
        client.run(os.environ["DISCORD_TOKEN"])
