from __future__ import annotations

import os
import re
import time
import json
from datetime import datetime
from typing import Any, Dict, List, Optional
from pathlib import Path

import requests


GAMMA_API_BASE = "https://public-api.gamma.app/v1.0"
def _save_checkpoint(state: dict) -> str:
    outdir = Path("output") / "checkpoints"
    outdir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = outdir / f"deck_checkpoint_{ts}.json"
    with open(path, "w", encoding="utf-8") as f:
        json.dump(state.get("deck_json", {}), f, ensure_ascii=False, indent=2)
    print(f"[CHECKPOINT] deck_json saved: {path}")
    return str(path)


def _slides_to_input_text(deck: Dict[str, Any]) -> str:
    title = (deck.get("deck_title") or "").strip() or "발표자료"
    slides: List[Dict[str, Any]] = deck.get("slides") or []
    n = len(slides)

    header = f"""[DECK]
DECK_TITLE: {title}
TOTAL_SLIDES: {n}

ABSOLUTE RULES:
- 정확히 {n}장만 생성. 추가/삭제/분할/병합 금지.
- 슬라이드 순서 변경 금지.
- 사진/실사/캐릭터/배경 이미지 생성 금지.
- 빈 이미지 placeholder(회색 박스, 깨진 아이콘) 생성 금지.
- 텍스트는 한국어 중심으로 작성(고유명사/약어만 예외).
[/DECK]
""".strip()

    def _strip_formal_endings(text: str) -> str:
        # 종결어미 제거 비활성화: 문장 절단/어색한 마침표 방지
        return str(text or "").strip()

    def _clean_lines(xs: List[str], limit: int) -> List[str]:
        out: List[str] = []
        for x in xs:
            x = _strip_formal_endings(str(x or "")).strip()
            if not x:
                continue
            if x in {"**POST_DIAGRAM_SYSTEM**", "**POST_DIAGRAM_ORGCHART**", "**후처리_대상**"}:
                continue
            out.append(x)
            if len(out) >= limit:
                break
        return out

    slide_blocks: List[str] = []
    for i, s in enumerate(slides, 1):
        section = _strip_formal_endings((s.get("section") or "").strip())
        slide_title = _strip_formal_endings((s.get("slide_title") or "").strip()) or "슬라이드"
        key_message = _strip_formal_endings((s.get("key_message") or "").strip())

        bullets = _clean_lines(s.get("bullets") or [], limit=7)

        table_md = _strip_formal_endings((s.get("TABLE_MD") or "").strip())
        diagram_spec = _strip_formal_endings((s.get("DIAGRAM_SPEC_KO") or "").strip())
        chart_spec = _strip_formal_endings((s.get("CHART_SPEC_KO") or "").strip())

        evidence = s.get("evidence") or []
        ev_lines: List[str] = []
        if isinstance(evidence, list):
            for ev in evidence[:3]:
                if isinstance(ev, dict):
                    t = (ev.get("type") or "근거").strip()
                    tx = _strip_formal_endings((ev.get("text") or "").strip())
                    if tx:
                        ev_lines.append(f"- ({t}) {tx}")
                else:
                    tx = _strip_formal_endings(str(ev or "").strip())
                    if tx:
                        ev_lines.append(f"- {tx}")

        lines: List[str] = []
        lines.append(f"[SLIDE {i}/{n}]")
        lines.append(f"SECTION: {section}")
        lines.append(f"TITLE: {slide_title}")
        if key_message:
            lines.append(f"KEY_MESSAGE: {key_message}")
        lines.append(f"SLIDE_LAYOUT: {(s.get('slide_layout') or '').strip()}")
        lines.append(f"VISUAL_SLOT: {(s.get('visual_slot') or '').strip()}")
        lines.append(f"CONTENT_DENSITY: {(s.get('content_density') or '').strip()}")
        lines.append(f"IMAGE_NEEDED: {bool(s.get('image_needed'))}")
        lines.append(f"IMAGE_TYPE: {(s.get('image_type') or 'none')}")
        if str(s.get("image_brief_ko") or "").strip():
            lines.append(f"IMAGE_BRIEF_KO: {str(s.get('image_brief_ko') or '').strip()}")

        lines.append("BULLETS:")
        if bullets:
            for b in bullets:
                lines.append(f"- {b}")

        if ev_lines:
            lines.append("EVIDENCE:")
            lines.extend(ev_lines)

        if table_md:
            lines.append("TABLE_MD:")
            lines.append(table_md)
        if diagram_spec:
            lines.append("DIAGRAM_SPEC_KO:")
            lines.append(diagram_spec)
        if chart_spec:
            lines.append("CHART_SPEC_KO:")
            lines.append(chart_spec)

        lines.append("[ENDSLIDE]")
        slide_blocks.append("\n".join(lines))

    body = "\n\n---\n\n".join(slide_blocks)
    return header + "\n\n" + body


def _gamma_headers(api_key: str) -> Dict[str, str]:
    return {"X-API-KEY": api_key, "Content-Type": "application/json"}


def _list_themes(
    api_key: str,
    *,
    query: str = "",
    limit: int = 50,
    max_pages: int = 5,
) -> List[Dict[str, Any]]:
    themes: List[Dict[str, Any]] = []
    after = ""
    for _ in range(max_pages):
        params: Dict[str, Any] = {"limit": int(limit)}
        if query:
            params["query"] = query
        if after:
            params["after"] = after
        r = requests.get(
            f"{GAMMA_API_BASE}/themes",
            headers=_gamma_headers(api_key),
            params=params,
            timeout=60,
        )
        if r.status_code != 200:
            raise RuntimeError(f"Gamma themes API error {r.status_code}: {r.text}")
        payload = r.json() or {}
        data = payload.get("data") or []
        if isinstance(data, list):
            themes.extend([x for x in data if isinstance(x, dict)])
        if not payload.get("hasMore"):
            break
        after = str(payload.get("nextCursor") or "").strip()
        if not after:
            break
    return themes


def _resolve_theme_id(api_key: str, theme_input: Optional[str]) -> Optional[str]:
    raw = str(theme_input or "").strip()
    if not raw:
        return None
    # if id passed directly
    if re.fullmatch(r"[A-Za-z0-9_-]{8,}", raw):
        return raw

    themes = _list_themes(api_key, query=raw, limit=50, max_pages=5)
    if not themes:
        return None

    for t in themes:
        if str(t.get("name") or "").strip().lower() == raw.lower():
            return str(t.get("id") or "").strip() or None
    return str((themes[0] or {}).get("id") or "").strip() or None


def _start_generation(
    api_key: str,
    *,
    input_text: str,
    theme_id: Optional[str],
    num_cards: int,
) -> Dict[str, Any]:

    payload: Dict[str, Any] = {
        "inputText": input_text,
        "format": "presentation",
        "exportAs": "pptx",
        "textMode": "preserve",
        "numCards": int(num_cards),
        "cardOptions": {"dimensions": "16x9"},
        "cardSplit": "inputTextBreaks",

        # ???듭떖: ?대?吏 ?꾩쟾 李⑤떒(吏?쒕Ц?쇰줈??紐?留됰뒗 寃쎌슦媛 留롮쓬)
        "imageOptions": {"source": "noImages"},

        "textOptions": {
            "language": "ko",
            "tone": "professional, clear",
            "amount": "medium",
        },

        "additionalInstructions": (
            f"정확히 {int(num_cards)}장만 생성. 추가/삭제/분할/병합 금지.\n"
            f"슬라이드 순서 변경 금지.\n"
            f"SECTION 블록 순서 절대 유지: 기관 소개 -> 연구 개요 -> 연구 필요성 -> 연구 목표 -> 연구 내용 -> 추진 계획 -> 활용방안 및 기대효과 -> 사업화 전략 및 계획 -> Q&A.\n"
            f"한 섹션이 시작되면 다음 섹션으로 넘어가기 전까지 해당 섹션 슬라이드를 연속 배치.\n"
            f"영어 문장/영어 제목 금지(고유명사/약어만 예외).\n"
            f"사진/실사/캐릭터/배경 이미지 생성 금지.\n"
            f"TABLE_MD / CHART_SPEC_KO / DIAGRAM_SPEC_KO가 있으면 반드시 반영.\n"
            f"텍스트 밀도 과소 금지: 긴 문단은 금지하되, 슬라이드 당 정보 블록 최소 2개 이상 배치.\n"
            f"설명 문장보다 구조화된 정보 전달(표/도식) 우선.\n"
            f"'추가 정보/문의/연락처' 같은 마무리 슬라이드 생성 금지.\n"
            f"마지막은 '감사합니다' 1장만 허용(중복 금지).\n"
            f"디자인 스타일: 깔끔한 카드형 레이아웃, 균형 배치, 둥근 모서리 중심.\n"
            f"연한 회색 배경 + 블루 포인트 톤. 과도한 빈 공간 금지.\n"
            f"슬라이드별 SLIDE_LAYOUT / VISUAL_SLOT / CONTENT_DENSITY 힌트를 우선 적용.\n"
            f"NotebookLM 스타일처럼 제목-요약-구조화 정보 순서를 유지하고 카드 비율을 일정하게 배치.\n"
            f"한 슬라이드당 핵심 메시지 1개, 불릿은 3~5개 권장.\n"
            f"빈 공간이 크면 카드 2열/요약 박스/표/도식으로 반드시 채운다.\n"
            f"IMAGE_NEEDED=true 인 슬라이드는 빈 이미지 슬롯/회색 박스/깨진 이미지 아이콘을 절대 만들지 않는다.\n"
            f"IMAGE_NEEDED=true 인 슬라이드는 layout=text_image로 생성하고, 우측 40% 영역은 도형/다이어그램/표 등 실제 시각요소로 직접 채운다.\n"
            f"텍스트 박스/도형/표는 이미지/시각요소 영역을 침범하지 않도록 배치(겹침 금지).\n"
            f"시각요소를 생성할 수 없으면 해당 슬라이드를 만들지 말고 이전/다음 슬라이드에 내용 통합.\n"
            f"입력 블록의 TITLE/KEY_MESSAGE/BULLETS는 의미 변경 없이 최대한 원문 그대로 사용.\n"
            f"문장 축약, 재서술, 표현 치환 최소화. 특히 TITLE은 원문 유지.\n"
            f"SLIDE 블록 1개를 카드 1장으로 1:1 매핑하고, 블록 병합/분할 금지.\n"
            f"문장 형태로 작성하지 않는다. 모든 항목은 명사구 또는 키워드 형태로 작성.\n"
            f"문장 종결어미 사용 금지 (~다, ~니다, ~합니다, ~됩니다 포함).\n"
            f"발표 슬라이드용 bullet 형태로 작성. 최소 3개 bullet이 없으면 해당 슬라이드 생성 금지.\n"
            f"내용이 부족하면 슬라이드를 만들지 않는다.\n"
            f"목차 슬라이드에서는 제목만 출력하고 설명 문장은 출력하지 않는다.\n"
            f"표 생성 시 헤더 행 강조 색상, 행별 연한 alternating 색상, 글자색 대비 확보.\n"
            f"When a table is needed, use a simple table.\n"
            f"Use PowerPoint table object.\n"
            f"Do not use card layout.\n"
            f"Do not use infographic style.\n"
            f"No rounded cards.\n"
            f"For section '연구 개요', avoid single standalone diagram-only slide.\n"
            f"Use structured explanatory layout with 2~3 boxes/cards and supporting bullets.\n"
            f"Prefer box/card comparison or matrix style that explains concept, scope, and context.\n"
            f"Keep enough explanatory text in '연구 개요' while maintaining concise bullet style.\n"
            f"Keep [SLIDE i/N] ... [ENDSLIDE] blocks unchanged."
        ),
    }

    # ?좑툘 theme??洹몃┝??源붿븘踰꾨━??寃쎌슦媛 留롮븘?? 湲곕낯? ?ъ슜 ????
    # ?ъ슜?먭? ?뺣쭚 ?먰븷 ?뚮쭔 state["gamma_theme_allow"]=True濡?耳쒕룄濡?
    if theme_id:
        payload["themeId"] = theme_id

    url = f"{GAMMA_API_BASE}/generations"
    r = requests.post(url, headers=_gamma_headers(api_key), json=payload, timeout=60)
    if r.status_code not in (200, 201):
        raise RuntimeError(f"Gamma API error {r.status_code}: {r.text}")
    return r.json()


def _poll_generation(api_key: str, generation_id: str, *, timeout_sec: int) -> Dict[str, Any]:
    t0 = time.time()
    last: Dict[str, Any] = {}
    while time.time() - t0 < timeout_sec:
        r = requests.get(f"{GAMMA_API_BASE}/generations/{generation_id}", headers=_gamma_headers(api_key), timeout=60)
        r.raise_for_status()
        last = r.json()

        status = (last.get("status") or "").lower()
        if status in {"completed", "complete", "succeeded", "success"}:
            return last
        if status in {"failed", "error"}:
            raise RuntimeError(f"Gamma generation failed: {last}")

        time.sleep(3)

    raise TimeoutError(f"Gamma generation polling timeout ({timeout_sec}s). last={last}")


def _download_file(url: str, out_path: str) -> None:
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with requests.get(url, stream=True, timeout=300) as r:
        r.raise_for_status()
        with open(out_path, "wb") as f:
            for chunk in r.iter_content(chunk_size=1024 * 1024):
                if chunk:
                    f.write(chunk)


def _avoid_windows_lock(path: str) -> str:
    base, ext = os.path.splitext(path)
    if not os.path.exists(path):
        return path
    for i in range(1, 200):
        cand = f"{base} ({i}){ext}"
        if not os.path.exists(cand):
            return cand
    return f"{base}_{int(time.time())}{ext}"


def _safe_filename(name: str) -> str:
    name = re.sub(r"[\\/:*?\"<>|]+", " ", str(name or ""))
    name = re.sub(r"\s+", " ", name).strip()
    if not name:
        return "result"
    # 遺덊븘???묐?/湲고샇 ?뺣━ ???⑥뼱 寃쎄퀎 湲곗??쇰줈 異뺤빟
    name = re.sub(r"[()\\[\\]{}]", "", name)
    max_len = 36
    if len(name) <= max_len:
        return name
    cut = name[:max_len + 1]
    ws = cut.rfind(" ")
    if ws >= 16:
        return cut[:ws].rstrip()
    return name[:max_len].rstrip()


def gamma_generation_node(state: Dict[str, Any]) -> Dict[str, Any]:
    api_key = os.environ.get("GAMMA_API_KEY")
    if not api_key:
        raise RuntimeError("GAMMA_API_KEY媛 ?놁뒿?덈떎. .env ?먮뒗 ?섍꼍蹂?섏뿉 ?ㅼ젙?섏꽭??")

    deck = state.get("deck_json") or {}
    slides = deck.get("slides") or []
    if not slides:
        raise RuntimeError("deck_json.slides媛 鍮꾩뼱?덉뒿?덈떎. merge_deck_node 寃곌낵瑜??뺤씤?섏꽭??")

    input_text = _slides_to_input_text(deck)

    output_dir = (state.get("output_dir") or "output").strip()

    # Default stable filename unless caller overrides.
    if not (state.get("output_filename") or "").strip():
        output_filename = "RanDi_발표자료.pptx"
    else:
        output_filename = (state.get("output_filename") or "").strip()

    out_path = _avoid_windows_lock(os.path.join(output_dir, output_filename))

    timeout_sec = int(state.get("gamma_timeout_sec") or 600)
    theme_input = (state.get("gamma_theme_id") or state.get("gamma_theme") or "").strip() or None
    theme_id = _resolve_theme_id(api_key, theme_input)
    if theme_input and not theme_id:
        print(f"[WARN] Gamma theme not found: {theme_input} (proceeding without themeId)")
    elif theme_id:
        print(f"[INFO] Gamma themeId resolved: {theme_id}")

    if state.get("save_checkpoint", False):
        _save_checkpoint(state)


    gen = _start_generation(api_key, input_text=input_text, theme_id=theme_id, num_cards=len(slides))
    generation_id = gen.get("generationId") or gen.get("id")
    if not generation_id:
        raise RuntimeError(f"Gamma ?묐떟??generationId媛 ?놁뒿?덈떎: {gen}")

    done = _poll_generation(api_key, generation_id, timeout_sec=timeout_sec)

    def _extract_url(d: Dict[str, Any]) -> str:
        return (
            d.get("exportUrl")
            or d.get("pptxUrl")
            or (d.get("exports") or {}).get("pptx")
            or ""
        )

    file_url = _extract_url(done)

    # ??completed 吏곹썑??URL????쾶 遺숇뒗 耳?댁뒪 ???理쒕? 45珥?
    if not file_url:
        t1 = time.time()
        while time.time() - t1 < 45:
            time.sleep(2.5)
            r = requests.get(f"{GAMMA_API_BASE}/generations/{generation_id}", headers=_gamma_headers(api_key), timeout=60)
            r.raise_for_status()
            done2 = r.json()
            file_url = _extract_url(done2)
            if file_url:
                done = done2
                break

    if not file_url:
        raise RuntimeError(f"Gamma ?꾨즺 ?묐떟???ㅼ슫濡쒕뱶 URL???놁뒿?덈떎: {done}")

    _download_file(file_url, out_path)

    state["final_ppt_path"] = out_path
    state["gamma_ppt_path"] = out_path
    return state
