"""화면 그리기(render_*) · 학생 링크 화면 · 생성 흐름(AI 호출 → 파일 → 화면 상태)

app.py에서 나눈 모듈입니다. 아래 층(dogae.base, dogae.ai, dogae.storage, dogae.exam_data, dogae.questions, dogae.word_docs)만 가져다 씁니다."""
import sys
import io
import streamlit as st
import ast
import hmac
import math
import shutil
import tempfile
import time
import pymupdf
import os
import re
import difflib
import sqlite3
import json
import uuid
import datetime
import requests
import zipfile
import pandas as pd
from google import genai
from google.genai import types
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import hashlib as _hashlib
import secrets as _pysecrets
import wave as _wave
from dogae.base import *  # noqa: F401,F403
from dogae.ai import *  # noqa: F401,F403
from dogae.storage import *  # noqa: F401,F403
from dogae.exam_data import *  # noqa: F401,F403
from dogae.questions import *  # noqa: F401,F403
from dogae.word_docs import *  # noqa: F401,F403


def render_interview_profile(profile, uni_name, major, interview_type, real_cases=None, preview=None):
    levels = (profile or {}).get("levels") or []
    primary = (profile or {}).get("primary")
    is_sangbu = "생기부" in interview_type
    if not levels:
        st.caption(f"📊 기출 DB에 '{uni_name}' 또는 '{major}' 관련 기출이 없어 출제 분석을 표시할 수 없습니다.")
        return
    head = f"{primary['label']} 기출 {primary['n']}건 기준" if primary else "표본이 적어 참고만 가능"
    with st.expander(f"📊 {uni_name} {major} 실제 출제 분석 — {head}", expanded=False):
        st.caption(
            "수험생이 복기한 면접 후기를 키워드 규칙으로 분류한 근사치입니다. "
            "건수가 적은 범위는 '경향'으로만 참고하세요."
        )
        mode = (preview or {}).get("mode")
        by_scope = {lv["scope"]: lv for lv in levels}
        n_uni = (profile or {}).get("n_uni", 0)
        nat = by_scope.get("전국 같은 학과")
        if mode == "exact":
            st.success(
                f"✅ **{uni_name}에 {major} 기출이 있습니다.** 이 학과 기출 {by_scope['대학+학과']['n'] if '대학+학과' in by_scope else preview['n_same_uni']}건을 "
                "기준으로 삼아 문항을 만듭니다."
            )
        elif mode == "style_transfer":
            srcs = preview.get("sources") or []
            st.warning(
                f"🔀 **{uni_name}에는 {major} 기출이 없습니다.** 그래서 두 자료를 결합해 {uni_name} 스타일로 재구성합니다.\n\n"
                f"- **질문 방식** ← {uni_name}의 다른 학과 기출 {n_uni:,}건을 비교 "
                f"({', '.join(preview.get('style_depts', [])[:6]) or '다른 학과'} 등)\n"
                f"- **전공 내용** ← 다른 대학의 {major} 기출"
                + (f" {nat['n']:,}건 중 선별" if nat else " 중 선별")
                + (f" ({', '.join(srcs[:6])}{' 등' if len(srcs) > 6 else ''})" if srcs else "")
                + (f"\n- 참고: {uni_name}의 유사 학과({', '.join(preview['same_uni_depts'][:4])}) 기출도 함께 봅니다."
                   if preview.get("same_uni_depts") else "")
            )
        elif mode == "uni_only":
            st.warning(
                f"🏫 **{major}와 비슷한 학과 기출이 전국 어디에도 없습니다.** {uni_name}의 다른 학과 기출 {n_uni:,}건에서 "
                "질문 방식만 가져오고, 전공 내용은 고교 교육과정 연계 개념으로 직접 설계합니다."
            )
        elif mode == "dept_only":
            st.warning(
                f"📘 **{uni_name}의 기출이 DB에 없습니다.** 다른 대학의 {major} 기출을 기준으로 만듭니다. "
                "이 대학만의 질문 방식은 반영되지 않습니다."
            )
        elif mode == "none":
            st.warning("참고할 기출이 없어 생기부 내용과 전공 개념만으로 문항을 만듭니다.")

        if is_sangbu and primary and primary["jesimun_pct"] >= 40:
            st.warning(
                f"⚠️ {primary['label']} 기출의 {primary['jesimun_pct']}%가 제시문·구술 문제입니다. "
                "이 대학·전형은 **'상위권 대학 제시문 기반 면접'** 방식으로 준비하는 편이 맞습니다."
            )
        alloc, basis = allocate_question_types(profile)
        if is_sangbu and alloc:
            st.info(
                "**이번에 만들 5세트의 유형 배분** — " + " · ".join(f"{t} {k}세트" for t, k in alloc)
                + f"\n\n기준: {basis}"
            )
        if not is_sangbu:
            if real_cases:
                st.info(
                    "**먼저 제시할 실제 기출 사례** (각 세트는 이 기출을 응용해 만듭니다)\n\n"
                    + "\n".join(
                        f"{i}. {c['대학']} · {c['학과']} · {c['전형']} — {_CASE_TIER_LABEL.get(c['tier'], '')}"
                        for i, c in enumerate(real_cases, 1)
                    )
                )
            else:
                st.warning("제시문·구술 기출이 DB에 없어 실제 사례 없이 창작 문제만 만듭니다.")

        tabs = st.tabs([f"{lv['scope']} · {lv['n']}건" for lv in levels])
        for tab, lv in zip(tabs, levels):
            with tab:
                st.markdown(
                    f"**{lv['label']}** — 꼬리질문 약 {lv['follow_pct']}% · 평균 질문 길이 약 {lv['avg_len']}자"
                )
                st.dataframe(
                    pd.DataFrame([{"질문 유형": t, "건수": c, "비중": pct} for t, c, pct in lv["dist"]]),
                    hide_index=True, width="stretch",
                    column_config={"비중": st.column_config.ProgressColumn("비중", min_value=0, max_value=100, format="%d%%")},
                )
                if lv["tracks"]:
                    st.caption("전형별 건수: " + " · ".join(
                        f"{tk} {n}건" + (f"(제시문·구술 {jp}%)" if jp >= 10 else "")
                        for tk, n, jp in lv["tracks"][:6]
                    ))
                for t, _, _ in lv["dist"][:3]:
                    qs = lv["samples"].get(t)
                    if qs:
                        st.markdown(f"**{t} 실제 질문**\n" + "\n".join(f"- {_md_safe(q)}" for q in qs))
        depts = profile.get("dept_list") or []
        if depts:
            shown = ", ".join(f"{d}({c})" for d, c in depts[:30])
            more = f" 외 {len(depts) - 30}개" if len(depts) > 30 else ""
            st.caption(f"이 대학 기출 보유 학과({len(depts)}개): {shown}{more}")

def logout():
    out_dir = st.session_state.get("_out_dir")
    if out_dir:
        shutil.rmtree(out_dir, ignore_errors=True)
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()

def require_login():
    """로그인이 켜져 있는데 아직 로그인하지 않았다면 로그인 화면만 보여주고 멈춥니다."""
    accounts, _ = _auth_config()
    if not accounts:
        return
    if st.session_state.get("auth_teacher") in accounts:
        return
    st.session_state.pop("auth_teacher", None)

    st.title("🎓 도개고 면접 마스터")
    st.caption("교내 시범운영 중입니다. 배정받은 아이디와 비밀번호로 로그인해 주세요.")

    locked_until = st.session_state.get("_login_locked_until", 0)
    if time.time() < locked_until:
        st.error(f"로그인 시도가 너무 많습니다. {int(locked_until - time.time()) + 1}초 뒤에 다시 시도해 주세요.")
        st.stop()

    with st.form("login_form"):
        teacher_id = st.text_input("아이디")
        password = st.text_input("비밀번호", type="password")
        submitted = st.form_submit_button("로그인", width="stretch")

    if submitted:
        teacher_id = (teacher_id or "").strip()
        expected = accounts.get(teacher_id)
        ok = expected is not None and hmac.compare_digest(
            (password or "").encode("utf-8"), expected.encode("utf-8")
        )
        if ok:
            st.session_state["auth_teacher"] = teacher_id
            st.session_state["_login_fails"] = 0
            st.rerun()
        fails = st.session_state.get("_login_fails", 0) + 1
        st.session_state["_login_fails"] = fails
        if fails >= 5:
            st.session_state["_login_locked_until"] = time.time() + 60
            st.session_state["_login_fails"] = 0
        time.sleep(1)
        st.error("아이디 또는 비밀번호가 올바르지 않습니다.")
    st.stop()

def render_save_status(prefix=""):
    """마지막 저장 시각과 저장 위치를 한 줄로 보여줍니다."""
    last = st.session_state.get("last_save_info")
    if not last:
        st.caption(f"{prefix}아직 이번 세션에서 저장된 기록이 없습니다.")
        return
    if last["ok"] and "구글시트" in last["where"]:
        st.caption(f"{prefix}☁️ 마지막 저장: {last['at']} · {last['where']}")
    elif last["ok"]:
        st.caption(f"{prefix}💾 마지막 저장: {last['at']} · {last['where']} (재배포 시 사라질 수 있음)")
    else:
        st.caption(f"{prefix}⚠️ 마지막 저장 실패: {last['at']} · 로컬에만 임시 저장됨")

def set_highlight_state(hl, raw_items, student_name, target_desc):
    st.session_state.highlight_data = hl
    st.session_state.highlight_raw_items = raw_items
    st.session_state.highlight_error = ""
    st.session_state.highlight_files = (
        create_highlight_word(hl, student_name, target_desc, is_teacher=False),
        create_highlight_word(hl, student_name, target_desc, is_teacher=True),
        create_qa_sheet_word(hl, student_name, target_desc),
    )

def add_more_questions(card_no, uni_name, major, profile, student_name, target_desc, api_key):
    """한 카드(문장)에 새 질문 2개를 붙입니다. 성공하면 True."""
    hl = st.session_state.get("highlight_data")
    raw_items = list(st.session_state.get("highlight_raw_items") or [])
    card = next((c for c in (hl or {}).get("items", []) if c["no"] == card_no), None)
    if not card or card["raw_idx"] >= len(raw_items):
        return False
    pa = hl["paras"][card["para"]]
    prompt = build_more_questions_prompt(card, hl["text"][pa["start"]:pa["end"]], uni_name, major,
                                         highlight_priority_text(profile, uni_name))
    try:
        raw = _call_gemini_retry(prompt, api_key, json_mode=True)
    except Exception as e:
        st.session_state.highlight_notice = "질문을 더 만들지 못했습니다: " + _friendly_error(e)
        return False
    new_qs = []
    for o in _extract_json_objects(raw):
        if isinstance(o.get("질문"), list):
            new_qs.extend(o["질문"])
        elif isinstance(o.get("질문"), str) and "구절" not in o:
            new_qs.append(o)
    if not new_qs:
        st.session_state.highlight_notice = "AI 응답에서 새 질문을 읽지 못했습니다. 다시 눌러 주세요."
        return False
    item = dict(raw_items[card["raw_idx"]])
    old = item.get("질문")
    if isinstance(old, str):
        old = [{"질문": old, "가이드": item.get("준비", "")}]
    item["질문"] = list(old or []) + new_qs
    raw_items[card["raw_idx"]] = item
    set_highlight_state(build_highlight_data(st.session_state.get("loaded_student_record_text", ""), raw_items),
                        raw_items, student_name, target_desc)
    return True

def add_answers(qkeys, uni_name, major, student_name, target_desc, api_key):
    """모범답안이 없는 질문에 모범답안을 만들어 붙입니다. 만든 개수를 돌려줍니다."""
    hl = st.session_state.get("highlight_data") or {}
    raw_items = list(st.session_state.get("highlight_raw_items") or [])
    want = set(qkeys)
    rows = []
    for c in hl.get("items", []):
        pa = hl["paras"][c["para"]]
        for q in c["qs"]:
            if q["key"] in want and not answer_of(q):
                rows.append({"id": q["key"], "q": q["q"], "quote": c["quote"], "guide": q.get("guide", ""),
                             "para": hl["text"][pa["start"]:pa["end"]]})
    if not rows:
        return 0
    got = {}
    try:
        for i in range(0, len(rows), 10):          # 한 번에 10개씩 (응답이 잘리지 않게)
            raw = _call_gemini_retry(build_answers_prompt(rows[i:i + 10], uni_name, major), api_key, json_mode=True)
            for o in _extract_json_objects(raw):
                for a in (o.get("answers") if isinstance(o.get("answers"), list) else [o]):
                    if isinstance(a, dict) and a.get("id") and str(a.get("모범답안", "")).strip():
                        got[str(a["id"]).strip()] = str(a["모범답안"]).strip()
    except Exception as e:
        st.session_state.highlight_notice = "모범답안을 만들지 못했습니다: " + _friendly_error(e)
    got = {k: v for k, v in got.items() if k in want}
    if not got:
        if not st.session_state.get("highlight_notice"):
            st.session_state.highlight_notice = "AI 응답에서 모범답안을 읽지 못했습니다. 다시 눌러 주세요."
        return 0
    idx = next((i for i, it in enumerate(raw_items)
                if isinstance(it, dict) and isinstance(it.get("모범답안"), dict) and "구절" not in it), None)
    if idx is None:
        raw_items.append({"모범답안": got})
    else:
        merged = dict(raw_items[idx]["모범답안"])
        merged.update(got)
        raw_items[idx] = {"모범답안": merged}
    set_highlight_state(build_highlight_data(st.session_state.get("loaded_student_record_text", ""), raw_items),
                        raw_items, student_name, target_desc)
    return len(got)

def add_gap_cards(uni_name, profile, student_name, target_desc, api_key):
    """질문이 하나도 없는 영역에 카드를 더 만듭니다. 성공하면 True."""
    hl = st.session_state.get("highlight_data") or {}
    areas = hl.get("gap_areas") or []
    rec = st.session_state.get("loaded_student_record_text", "")
    if not areas or not rec.strip():
        return False
    prompt = build_gap_cards_prompt(rec, areas, [c["quote"] for c in hl.get("items", [])],
                                    highlight_priority_text(profile, uni_name))
    try:
        raw = _call_gemini_retry(prompt, api_key, json_mode=True)
    except Exception as e:
        st.session_state.highlight_notice = "빈 영역 질문을 만들지 못했습니다: " + _friendly_error(e)
        return False
    new_items = [o for o in _extract_json_objects(raw) if str(o.get("구절", "")).strip()]
    raw_items = list(st.session_state.get("highlight_raw_items") or []) + new_items
    new_hl = build_highlight_data(rec, raw_items)
    if len(new_hl["items"]) <= len(hl.get("items", [])):
        st.session_state.highlight_notice = "AI가 고른 문장이 원문과 맞지 않아 새 카드를 만들지 못했습니다. 다시 눌러 주세요."
        return False
    set_highlight_state(new_hl, raw_items, student_name, target_desc)
    return True

def _render_question_only(hl, show, only_weak, checks, save_args, uni, major, student_name, target_desc, api_key):
    """형광펜에서 나온 질문만 모아 보여 주고, 모범답안은 눌러야 보이게 합니다."""
    all_rows = [(c, q) for c in hl["items"] for q in c["qs"]]
    rows = [(c, q) for c, q in all_rows if show[c["comp"]]]
    if only_weak:
        rows = [(c, q) for c, q in rows if checks.get(q["key"]) in ("애매", "모름")]
    missing = [q["key"] for _, q in all_rows if not answer_of(q)]
    st.caption(f"질문 {len(rows)}개 — 모범답안은 '✅ 모범답안 보기'를 눌러야 보입니다. 먼저 소리 내어 답해 보세요.")
    b1, b2 = st.columns([3, 2], vertical_alignment="center")
    with b1:
        if missing and st.button(f"✅ 모범답안 없는 {len(missing)}개 한 번에 만들기", key="hl_ans_all",
                                 width="stretch", disabled=not api_key):
            with st.spinner(f"모범답안 {len(missing)}개를 만드는 중입니다... (질문 10개당 30초 안팎)"):
                if add_answers(missing, uni, major, student_name, target_desc, api_key):
                    _quiet_save(save_args)
            st.rerun()
    with b2:
        open_all = st.toggle("모범답안 모두 펼치기", key="hl_open_all")
    with st.container(height=900, border=True):
        if not rows:
            st.info("보여줄 질문이 없습니다. 오른쪽에서 역량 선택이나 '애매·모름 질문만 보기'를 바꿔 보세요.")
        last_area = None
        for c, q in rows:
            area = hl["paras"][c["para"]]["area"] or c["area_key"]
            if area != last_area:
                st.markdown(f"##### 📂 {area}")
                last_area = area
            with st.container(border=True):
                pressure = " · ⚠️ 압박 대비" if q["type"] == "압박" else ""
                st.markdown(
                    f"**{q['no']}.** {_md_safe(q['q'])}  \n"
                    f":gray[{HL_COMPS[c['comp']]['icon']} {_q_tags(c, q)}{pressure} · 근거 {_marker(c['no'])} “{_md_safe(c['quote'])}”]"
                )
                wkey = f"chk_{q['key']}"
                cur = checks.get(q["key"], "미확인")
                st.radio(
                    "준비 상태", _CHECK_OPTS, horizontal=True, label_visibility="collapsed", key=wkey,
                    index=_CHECK_OPTS.index(cur) if cur in _CHECK_OPTS else 0,
                    on_change=_save_card_check, args=(q["key"], wkey, save_args),
                )
                with st.expander("✅ 모범답안 보기", expanded=open_all):
                    ans = answer_of(q)
                    if ans:
                        st.markdown(_md_safe(ans))
                    else:
                        if q.get("guide"):
                            st.caption("답변 가이드: " + q["guide"])
                        if st.button("이 질문 모범답안 만들기", key=f"hl_ans_{q['key']}", disabled=not api_key):
                            with st.spinner("모범답안을 만드는 중입니다..."):
                                if add_answers([q["key"]], uni, major, student_name, target_desc, api_key):
                                    _quiet_save(save_args)
                            st.rerun()
                    tail = (q.get("detail") or {}).get("꼬리질문")
                    if tail:
                        st.markdown(f"**🔥 꼬리질문**  \n{_md_safe(tail)}")

def render_card_view(hl, student_name, uni, major, interview_type, difficulty, target_desc, api_key):
    """형광펜 생기부 · 질문 카드 화면. 질문은 근거 문장 바로 아래에 붙습니다."""
    checks = st.session_state.setdefault("card_checks", {})
    save_args = (student_name, uni, major, interview_type, difficulty)
    profile = cached_interview_profile(APP_DIR, uni, major)

    result_text = st.session_state.get("last_result_text", "")
    unmatched_sets = attach_set_details(hl, result_text)
    n_sets = len(sangbu_sets(result_text))

    st.markdown("### 🖍️ 형광펜 생기부 · 질문 카드")
    st.caption(
        f"칠한 문장 {len(hl['items'])}곳 · 질문 {hl['q_total']}개 — 질문은 근거 문장 바로 아래에 붙어 있습니다."
        + (f" 5세트 문항 {n_sets - len(unmatched_sets)}개도 근거 문장에 합쳐 두었습니다." if n_sets else "")
    )
    if st.session_state.get("highlight_notice"):
        st.warning(st.session_state.pop("highlight_notice"))

    brief = sangbu_briefing(result_text)
    if brief:
        with st.container(border=True):
            st.markdown("**🔍 생기부 분석 요약**")
            items = list(brief.items())
            strength = next(((k, v) for k, v in items if "강점" in k), items[0])
            st.markdown(f"**{strength[0]}**  \n{_md_safe(strength[1])}")
            rest = [(k, v) for k, v in items if k != strength[0]]
            for k, v in rest:
                st.markdown(f"**{k}** — {_md_safe(_first_sentence(v))}")
            if rest:
                with st.expander("약점 방어 · 공략 포인트 자세히 (카드의 〰️ 밑줄·압박 질문과 같은 내용)"):
                    for k, v in rest:
                        st.markdown(f"**{k}**  \n{_md_safe(v)}")

    qonly = st.radio(
        "보기 방식", ["🖍️ 생기부에서 보기", "❓ 질문만 보기 (모범답안)"], horizontal=True,
        key="hl_view", label_visibility="collapsed",
    ).startswith("❓")
    left, right = st.columns([7, 3], gap="medium")
    with right:
        with st.container(border=True):
            st.markdown("**칠한 역량**")
            show = {k: st.checkbox(f"{v['icon']} {v['label']} ({hl['count'][k]})", value=True, key=f"hl_show_{k}")
                    for k, v in HL_COMPS.items()}
            show_def = st.checkbox(f"{HL_DEFEND['icon']} {HL_DEFEND['label']} 밑줄 ({hl['defend']})", value=True, key="hl_show_def")
            only_marked = st.checkbox("칠한 문단만 보기", value=True, key="hl_only_marked")
            only_weak = st.checkbox("애매·모름 질문만 보기", value=False, key="hl_only_weak")

        with st.container(border=True):
            keys = hl["q_keys"]
            done = sum(1 for k in keys if checks.get(k) == "자신 있음")
            n_amb = sum(1 for k in keys if checks.get(k) == "애매")
            n_no = sum(1 for k in keys if checks.get(k) == "모름")
            st.markdown(f"**연습 진도** {done}/{len(keys)}")
            st.progress(done / len(keys) if keys else 0.0)
            st.caption(f"애매 {n_amb} · 모름 {n_no} — 질문마다 아래에서 표시하면 기록에 저장됩니다.")

        with st.container(border=True):
            st.markdown("**질문 더 만들기**")
            if hl["gap_areas"]:
                st.caption("질문이 하나도 없는 영역: " + ", ".join(hl["gap_areas"]))
                if st.button(f"➕ 질문 없는 {len(hl['gap_areas'])}곳에 질문 만들기", width="stretch",
                             key="hl_fill_gaps", disabled=not api_key):
                    with st.spinner("빈 영역에서 문장을 골라 질문을 만드는 중입니다..."):
                        if add_gap_cards(uni, profile, student_name, target_desc, api_key):
                            _quiet_save(save_args)
                    st.rerun()
            else:
                st.caption("모든 영역에 질문이 있습니다.")
            st.caption("문장마다 '➕ 질문 더 만들기'를 누르면 그 문장에서 새 질문이 2개씩 나옵니다.")
            if hl["gaps"]:
                with st.popover("영역별 준비 상태", width="stretch"):
                    st.dataframe(pd.DataFrame(hl["gaps"]), hide_index=True, width="stretch")
                    if hl.get("subjects"):
                        st.caption("질문이 닿은 세특 과목: " + ", ".join(hl["subjects"]))

        with st.container(border=True):
            st.markdown(f"**{uni} 면접 정보**")
            try:
                info = get_univ_info(RT.univ_info_db, uni, major)
            except Exception:
                info = None
            if info is not None:     # 대학·학과 한 줄(pandas Series)
                for col in ["면접유형", "면접시간", "면접위원"]:
                    val = str(info.get(col, "") or "").strip()
                    if val and val.lower() != "nan":
                        st.caption(f"{col}: {val}")
            share, basis = pick_allocation_basis(profile)
            if share:
                for t, v in sorted(share.items(), key=lambda kv: -kv[1])[:4]:
                    st.caption(f"{t} {int(round(v * 100))}%")
                st.caption(f"기준: {basis}")
            elif info is None:
                st.caption("이 대학의 면접 자료가 아직 없습니다.")

        if hl.get("keywords"):
            with st.container(border=True):
                st.markdown("**핵심 키워드**")
                st.markdown(" ".join(f"`{k}`" for k in hl["keywords"]))

    with left:
        if qonly:
            _render_question_only(hl, show, only_weak, checks, save_args, uni, major, student_name, target_desc, api_key)
        else:
            with st.container(height=900, border=True):
                text, shown_any, last_area = hl["text"], False, None
                for pi, pa in enumerate(hl["paras"]):
                    start, end = pa["start"], pa["end"]
                    if start == end:
                        continue
                    cards = [c for c in hl["items"] if c["para"] == pi and show[c["comp"]]]
                    if only_weak:
                        cards = [c for c in cards if any(checks.get(q["key"]) in ("애매", "모름") for q in c["qs"])]
                    if not cards and (only_marked or only_weak):
                        continue
                    shown_any = True
                    if pa["area"] and pa["area"] != last_area:
                        st.caption(f"📂 {pa['area']}")
                        last_area = pa["area"]
                    st.markdown(_para_html(text, start, end, cards, show_def), unsafe_allow_html=True)
                    for c in cards:
                        with st.container(border=True):
                            for q in c["qs"]:
                                if only_weak and checks.get(q["key"]) not in ("애매", "모름"):
                                    continue
                                pressure = " · ⚠️ 압박 대비" if q["type"] == "압박" else ""
                                st.markdown(
                                    f"**{q['no']}.** {_marker(c['no'])} {_md_safe(q['q'])}  \n"
                                    f":gray[{HL_COMPS[c['comp']]['icon']} {_q_tags(c, q)}{pressure}]"
                                )
                                chk_col, guide_col = st.columns([3, 1.4], vertical_alignment="center")
                                with chk_col:
                                    wkey = f"chk_{q['key']}"
                                    cur = checks.get(q["key"], "미확인")
                                    st.radio(
                                        "준비 상태", _CHECK_OPTS, horizontal=True, label_visibility="collapsed", key=wkey,
                                        index=_CHECK_OPTS.index(cur) if cur in _CHECK_OPTS else 0,
                                        on_change=_save_card_check, args=(q["key"], wkey, save_args),
                                    )
                                with guide_col:
                                    if q.get("guide") or q.get("detail") or q.get("answer"):
                                        with st.popover("답변 가이드", width="stretch"):
                                            if q.get("guide"):
                                                st.markdown(_md_safe(q["guide"]))
                                            if q.get("detail"):
                                                st.markdown(_set_detail_md(q["detail"]))
                                            elif q.get("answer"):
                                                st.markdown(f"**✅ 모범 답안**  \n{_md_safe(q['answer'])}")
                            if st.button("➕ 이 문장으로 질문 더 만들기", key=f"hl_more_{c['no']}_{c['raw_idx']}",
                                         disabled=not api_key):
                                with st.spinner("이 문장에서 새 질문을 만드는 중입니다..."):
                                    if add_more_questions(c["no"], uni, major, profile, student_name, target_desc, api_key):
                                        _quiet_save(save_args)
                                st.rerun()
                if not shown_any:
                    st.info("보여줄 문장이 없습니다. 오른쪽에서 역량 선택이나 '애매·모름 질문만 보기'를 바꿔 보세요.")
        if unmatched_sets:
            with st.expander(f"📌 근거 문장에 붙지 않은 5세트 문항 {len(unmatched_sets)}개"):
                for d in unmatched_sets:
                    st.markdown(f"**5세트 {d['no']}번 · {_md_safe(d['title'])}**  \n{_md_safe(d['질문'])}")
                    st.markdown(_set_detail_md({k: v for k, v in d.items() if k != "질문"}))
                    st.divider()
        if hl.get("dropped"):
            st.caption(f"AI가 고른 문장 중 {hl['dropped']}개는 원문과 글자가 달라 칠하지 않았습니다(원문을 바꾸지 않기 위해).")

def set_interview_card_state(card, student_name, target_desc, uni, major):
    st.session_state.interview_card = card
    st.session_state.interview_card_file = None
    if card:
        try:
            info = get_univ_info(RT.univ_info_db, uni, major)
        except Exception:
            info = None
        st.session_state.interview_card_file = create_interview_card_word(
            card, student_name, target_desc, info_row=info, keywords=_card_keywords())

def render_interview_card_section(student_name, uni, major, target_desc, api_key):
    st.markdown("### 🗂️ 면접 전날 최종 요약카드")
    st.caption("생기부 속 용어·개념의 뜻과 출처(학년·과목·내용), 답변에 덧붙일 말, "
               f"{major or '지원 학과'} 예상 질문 {CARD_Q_TOTAL}선과 모범답안을 한 문서로 묶습니다.")
    card = st.session_state.get("interview_card")
    rec = st.session_state.get("loaded_student_record_text", "")
    if card and (card.get("uni"), card.get("major")) != (uni, major):
        st.info(f"이 카드는 {card.get('uni')} {card.get('major')} 기준으로 만든 것입니다. 지원 대학·학과를 바꿨다면 다시 만들어 주세요.")
    b1, b2 = st.columns([2, 3], vertical_alignment="center")
    with b1:
        label = "🔄 요약카드 다시 만들기" if card else "🗂️ 요약카드 만들기"
        go = st.button(label, width="stretch", key="make_interview_card",
                       disabled=not api_key or not rec.strip() or not major)
    with b2:
        if not rec.strip():
            st.caption("생기부 원문이 없습니다. 생기부 PDF를 올리고 면접 패키지를 먼저 만들어 주세요.")
        elif not card:
            st.caption("AI를 두 번 불러 만듭니다(1~2분).")
    if go:
        with st.spinner(f"핵심 개념과 {major} 예상 질문 {CARD_Q_TOTAL}선을 만드는 중입니다... (1~2분)"):
            new_card, err = make_interview_card(
                rec, uni, major, student_name, cached_interview_profile(APP_DIR, uni, major), RT.exam_db, api_key,
                bank=current_bank())
        st.session_state.interview_card_error = err
        if new_card:
            set_interview_card_state(new_card, student_name, target_desc, uni, major)
            save_interview_card(st.session_state.get("current_record_id"), student_name, uni, major, new_card)
        st.rerun()
    if st.session_state.get("interview_card_error"):
        st.warning(st.session_state.interview_card_error)
    if not card:
        return

    path = st.session_state.get("interview_card_file")
    if path and os.path.exists(path):
        with open(path, "rb") as f:
            st.download_button("🗂️ 면접 전날 요약카드 다운로드 (.docx)", f, file_name=os.path.basename(path),
                               width="stretch", key="dl_interview_card_inline")
    concepts, qas = card.get("concepts") or [], card.get("qas") or []
    tab_c, tab_q = st.tabs([f"🔤 핵심 개념 {len(concepts)}개", f"❓ 예상 질문 {len(qas)}선"])
    with tab_c:
        with st.container(height=700, border=False):
            for n, c in enumerate(concepts, 1):
                with st.container(border=True):
                    warn = " :red[⚠️ 원문 표기 확인]" if not c.get("found") else ""
                    kind = f" :gray[{c['kind']}]" if c.get("kind") else ""
                    st.markdown(f"**{n}. {_md_safe(c['term'])}**{kind}{warn}  \n{_md_safe(c.get('desc', ''))}")
                    if c.get("srcs"):
                        st.markdown("  \n".join(
                            f"📍 **{_md_safe(s['grade'])} · {_md_safe(s['area'])}** — {_md_safe(s.get('what', ''))}"
                            for s in c["srcs"]))
                    if c.get("add"):
                        st.markdown(f"➕ **덧붙일 말** {_md_safe(c['add'])}")
                    if c.get("follow"):
                        st.caption(f"🔥 꼬리질문: {c['follow']}")
    with tab_q:
        open_all = st.toggle("모범답안 모두 펼치기", key="card_open_all")
        with st.container(height=700, border=False):
            for n, qa in enumerate(qas, 1):
                with st.container(border=True):
                    meta = [qa.get("type", "")]
                    if qa.get("basis") and qa["basis"] != "없음":
                        meta.append("근거 " + qa["basis"])
                    if qa.get("ref"):
                        meta.append("기출 참고 " + qa["ref"])
                    st.markdown(f"**{n}. {_md_safe(qa['q'])}**  \n:gray[{_md_safe(' · '.join(m for m in meta if m))}]")
                    with st.expander("✅ 모범답안 보기", expanded=open_all):
                        st.markdown(_md_safe(qa["answer"]))
                        if qa.get("keys"):
                            st.caption("핵심: " + " · ".join(qa["keys"]))

def set_question_sheet_state(sheet, student_name, target_desc, uni, major):
    st.session_state.qsheet = sheet
    st.session_state.qsheet_files = None
    if sheet:
        try:
            info = get_univ_info(RT.univ_info_db, uni, major)
        except Exception:
            info = None
        st.session_state.qsheet_files = (
            create_question_sheet_word(sheet, student_name, target_desc, False, info),
            create_question_sheet_word(sheet, student_name, target_desc, True, info),
        )

def render_question_sheet_section(student_name, uni, major, target_desc, api_key, show_preview=True):
    st.markdown("#### 📋 대학 맞춤 면접 질문지 (AI)" if not show_preview else "### 📋 대학 맞춤 면접 질문지")
    st.caption(f"{uni or '지원 대학'}의 실제 질문 방식과 {major or '지원 학과'} 기출의 전공 깊이로, 이 학생 생기부에서 질문을 뽑습니다. "
               "학생용은 질문과 근거 생기부 문장만, 교사용은 모범답안·꼬리질문·피할 답변까지 넣습니다.")
    sheet = st.session_state.get("qsheet")
    rec = st.session_state.get("loaded_student_record_text", "")
    try:
        info = get_univ_info(RT.univ_info_db, uni, major)
    except Exception:
        info = None
    if sheet and (sheet.get("uni"), sheet.get("major")) != (uni, major):
        st.info(f"이 질문지는 {sheet.get('uni')} {sheet.get('major')} 기준입니다. 지원 대학·학과를 바꿨다면 다시 만들어 주세요.")
    b0, b1, b2 = st.columns([1.2, 2, 2.8], vertical_alignment="bottom")
    with b0:
        default_n = default_qsheet_count(info) if info is not None else 15
        n = st.selectbox("문항 수", QSHEET_COUNTS, index=QSHEET_COUNTS.index(default_n), key="qsheet_n",
                         help="지원 대학의 면접 시간에 맞춰 기본값을 정했습니다.")
    with b1:
        go = st.button("🔄 질문지 다시 만들기" if sheet else "📋 질문지 만들기", width="stretch", key="make_qsheet",
                       disabled=not api_key or not rec.strip() or not major)
    with b2:
        if not rec.strip():
            st.caption("생기부 원문이 없습니다. 생기부 PDF를 올리고 면접 패키지를 먼저 만들어 주세요.")
        elif not sheet:
            st.caption("AI를 한 번 불러 만듭니다(약 1분).")
    if go:
        with st.spinner(f"{uni} 질문 방식으로 {n}개 질문을 뽑는 중입니다... (약 1분)"):
            new_sheet, err = make_question_sheet(rec, uni, major, student_name, n, RT.exam_db, info, api_key,
                                                 bank=current_bank())
        st.session_state.qsheet_error = err
        if new_sheet:
            set_question_sheet_state(new_sheet, student_name, target_desc, uni, major)
            save_question_sheet(st.session_state.get("current_record_id"), student_name, uni, major, new_sheet)
        st.rerun()
    if st.session_state.get("qsheet_error"):
        st.warning(st.session_state.qsheet_error)
    if not sheet:
        return

    files = st.session_state.get("qsheet_files") or (None, None)
    d1, d2 = st.columns(2)
    for col, path, label, key in ((d1, files[0], "📘 질문지 학생용 (.docx)", "dl_qsheet_stu"),
                                  (d2, files[1], "📕 질문지 교사용 · 모범답안 (.docx)", "dl_qsheet_tea")):
        if path and os.path.exists(path):
            with open(path, "rb") as f:
                col.download_button(label, f, file_name=os.path.basename(path), width="stretch", key=key)
    st.caption(f"출제 기준: {sheet.get('basis', '')}" + (f" · {sheet['style_line']}" if sheet.get("style_line") else ""))
    if not show_preview:
        reused = sum(1 for q in sheet.get("qs") or [] if q.get("from"))
        st.caption(f"질문 {len(sheet.get('qs') or [])}개" + (f" (은행 질문을 다듬은 것 {reused}개)" if reused else "")
                   + " — 내용은 위 🏦 질문 은행 표에서 '질문지'로 보입니다.")
        return
    open_all = st.toggle("교사용 내용(모범답안·꼬리질문) 모두 펼치기", key="qsheet_open_all")
    with st.container(height=700, border=False):
        for n_, q in enumerate(sheet.get("qs") or [], 1):
            with st.container(border=True):
                tags = " · ".join(x for x in (q.get("type"), q.get("eval")) if x)
                st.markdown(f"**{n_}. {_md_safe(q['q'])}**  \n:gray[{_md_safe(tags)}]")
                if q.get("quote") or _src_line(q):
                    color = "red" if q.get("found") is False else "gray"
                    st.markdown(f"📍 **{_md_safe(_src_line(q) or '생기부')}**{_md_safe(_found_mark(q))}  \n"
                                f":{color}[“{_md_safe(q.get('quote', ''))}”]")
                with st.expander("📕 교사용: 모범답안 · 꼬리질문", expanded=open_all):
                    if q.get("style"):
                        st.caption("🏫 " + q["style"])
                    st.markdown(_md_safe(q.get("answer", "")))
                    if q.get("keys"):
                        st.caption("핵심: " + " · ".join(q["keys"]))
                    if q.get("follow"):
                        st.markdown(f"🔥 **꼬리질문** {_md_safe(q['follow'])}")
                    if q.get("avoid"):
                        st.caption("✖ 피할 답변: " + q["avoid"])


def current_bank(is_sb=True):
    """지금 학생의 질문 은행(같은 재료면 다시 계산하지 않음).
    제시문 면접은 그 면접의 문항만 씁니다(앞서 본 생기부 학생의 카드·질문지가 섞이지 않게)."""
    hl = (st.session_state.get("highlight_data") or {}) if is_sb else {}
    qs = (st.session_state.get("qsheet") or {}) if is_sb else {}
    cd = (st.session_state.get("interview_card") or {}) if is_sb else {}
    res = st.session_state.get("last_result_text", "") or ""
    rec = st.session_state.get("loaded_student_record_text", "") or ""
    sig = (is_sb, _hashlib.sha1(res.encode("utf-8")).hexdigest()[:12], tuple(hl.get("q_keys") or ()),
           len(hl.get("items") or []), qs.get("made"), len(qs.get("qs") or []), cd.get("made"), len(rec))
    cached = st.session_state.get("_qbank_cache")
    if cached and cached[0] == sig:
        return cached[1]
    try:
        bank = build_question_bank(hl, res, qs, cd, rec)
    except Exception:
        bank = []
    st.session_state["_qbank_cache"] = (sig, bank)
    return bank


def goto_section(name):
    """다음 화면 그리기 때 결과 화면을 바꿉니다(위젯이 이미 그려진 뒤에도 안전)."""
    st.session_state["_goto_section"] = name


def render_question_bank_section(student_name, uni, major, target_desc, api_key):
    bank = current_bank()
    st.markdown("### 🏦 질문 은행")
    st.caption("이 학생을 위해 만든 질문(질문지·형광펜·요약카드·5세트)을 한곳에 모았습니다. "
               "**쓸 질문을 고르면** 학생 휴대폰 링크 · 🎙️ 모의면접 진행 · 질문지 내보내기가 모두 이 선택을 씁니다.")
    if bank:
        sel = bank_selection(bank)
        chosen = set(sel)
        counts = {s: sum(1 for e in bank if e["src"] == s) for s in QB_SOURCES}
        dups = sum(1 for e in bank if e.get("dup_of"))
        st.caption(" · ".join(f"{s} {n}" for s, n in counts.items() if n)
                   + (f" · 비슷한 질문 {dups}개는 하나로 묶음" if dups else "") + f" · **고른 질문 {len(sel)}개**")
        present = [s for s in QB_SOURCES if counts[s]]
        f1, f2, f3 = st.columns([3, 2, 2], vertical_alignment="center")
        with f1:
            srcs = st.pills("출처", present, selection_mode="multi", default=present, key="qb_src",
                            label_visibility="collapsed") or present
        with f2:
            hide_dup = st.toggle("비슷한 질문 숨기기", value=True, key="qb_hide_dup")
        with f3:
            only_sel = st.toggle("고른 질문만", value=False, key="qb_only_sel")
        view = [e for e in bank if e["src"] in srcs
                and not (hide_dup and e.get("dup_of") and e["key"] not in chosen)
                and (not only_sel or e["key"] in chosen)]
        if view:
            df = pd.DataFrame([{
                "선택": e["key"] in chosen, "출처": f"{e['src']} {e['no']}",
                "질문": e["q"] + (f"  ← {e['based_on']}" if e.get("based_on") else ""),
                "근거": _qb_src_line(e) + (" ⚠️" if e.get("found") is False else ""),
                "답안": "✅" if e.get("answer") else "", "_key": e["key"]} for e in view])
            ver = st.session_state.get("qb_ver", 0)
            sig = _hashlib.sha1("|".join(e["key"] for e in view).encode()).hexdigest()[:8]
            edited = st.data_editor(
                df, key=f"qb_editor_{ver}_{sig}", hide_index=True, width="stretch",
                height=min(560, 40 + 36 * len(df)), disabled=["출처", "질문", "근거", "답안"],
                column_config={
                    "선택": st.column_config.CheckboxColumn("선택", width="small"),
                    "출처": st.column_config.TextColumn("출처", width="small"),
                    "질문": st.column_config.TextColumn("질문", width="large"),
                    "근거": st.column_config.TextColumn("생기부 근거", width="medium"),
                    "답안": st.column_config.TextColumn("답안", width="small"),
                    "_key": None,
                })
            shown = set(df["_key"])
            new_sel = (chosen - shown) | set(edited.loc[edited["선택"], "_key"])
            if new_sel != chosen:
                st.session_state["qbank_sel"] = [e["key"] for e in bank if e["key"] in new_sel]
                st.session_state["qb_ver"] = ver + 1
                st.rerun()
        else:
            st.info("이 보기에 해당하는 질문이 없습니다.")

        a1, a2, a3 = st.columns(3)
        if a1.button("☑️ 보이는 질문 모두 고르기", width="stretch", key="qb_all"):
            st.session_state["qbank_sel"] = [e["key"] for e in bank if e["key"] in chosen | {v["key"] for v in view}]
            st.session_state["qb_ver"] = st.session_state.get("qb_ver", 0) + 1
            st.rerun()
        if a2.button("⬜ 고른 것 모두 풀기", width="stretch", key="qb_none"):
            st.session_state["qbank_sel"] = ["__none__"]
            st.session_state["qb_ver"] = st.session_state.get("qb_ver", 0) + 1
            st.rerun()
        if a3.button("💾 고른 질문 저장", width="stretch", key="qb_save",
                     help="고른 질문은 모의면접 진행·학생 링크를 쓸 때도 자동으로 저장됩니다."):
            if save_qbank_state(st.session_state.get("current_record_id"), student_name, uni, major):
                st.toast("고른 질문을 저장했습니다.", icon="💾")

        picked = bank_selected_entries(bank)
        if picked:
            with st.container(border=True):
                st.markdown(f"**고른 질문 {len(picked)}개로**")
                b1, b2, b3, b4 = st.columns(4)
                try:
                    info = get_univ_info(RT.univ_info_db, uni, major)
                except Exception:
                    info = None
                sheet = bank_as_sheet(picked, uni, major)
                for col, teacher, label in ((b1, False, "📘 질문지 학생용"), (b2, True, "📕 질문지 교사용")):
                    path = create_question_sheet_word(sheet, student_name, target_desc, teacher, info, file_tag="은행")
                    with open(path, "rb") as f:
                        col.download_button(label, f, file_name=os.path.basename(path), width="stretch",
                                            key=f"qb_dl_{'t' if teacher else 's'}")
                if b3.button("🎙️ 모의면접 시작", type="primary", width="stretch", key="qb_run"):
                    save_qbank_state(st.session_state.get("current_record_id"), student_name, uni, major)
                    st.session_state["run"] = None
                    goto_section("🎙️ 모의면접 진행")
                    st.rerun()
                if b4.button("📱 학생 링크로", width="stretch", key="qb_link",
                             help="🎤 연습·첨삭 화면에서 학생 링크를 만들거나 '최신 질문 보내기'를 누르면 이 질문이 갑니다."):
                    save_qbank_state(st.session_state.get("current_record_id"), student_name, uni, major)
                    st.session_state["tp_src_pref"] = "bank"
                    goto_section("🎤 연습·첨삭")
                    st.rerun()

        with st.expander("🔍 질문 하나 자세히 보기 (모범답안 · 꼬리질문 · 근거 문단)"):
            pick = st.selectbox("질문", [e["key"] for e in bank], key="qb_detail",
                                format_func=lambda k: next(f"[{e['src']} {e['no']}] {e['q'][:60]}" for e in bank if e["key"] == k))
            e = next(x for x in bank if x["key"] == pick)
            meta = [x for x in (e.get("type"), e.get("eval"), _qb_src_line(e)) if x]
            st.markdown(f"**{_md_safe(e['q'])}**  \n:gray[{_md_safe(' · '.join(meta))}]")
            if e.get("dup_of"):
                tgt = next((x for x in bank if x["key"] == e["dup_of"]), None)
                if tgt:
                    st.caption(f"비슷한 질문: [{tgt['src']} {tgt['no']}] {tgt['q'][:80]}")
            if e.get("quote"):
                st.markdown(f"📍 “{_md_safe(e['quote'])}”")
            if e.get("para") and e.get("para") != e.get("quote"):
                st.caption(e["para"][:500])
            if e.get("style"):
                st.caption("🏫 " + e["style"])
            if e.get("answer"):
                st.markdown("**✅ 모범답안**  \n" + _md_safe(e["answer"]))
            if e.get("keys"):
                st.caption("핵심: " + " · ".join(e["keys"]))
            if e.get("tail"):
                st.markdown(f"🔥 **꼬리질문** {_md_safe(e['tail'])}")
            if e.get("avoid"):
                st.caption("✖ 피할 답변: " + e["avoid"])
    else:
        st.info("아직 질문이 없습니다. 아래에서 대학 맞춤 질문지를 만들거나, 형광펜 질문 카드를 만들어 주세요.")

    st.divider()
    with st.container(border=True):
        render_question_sheet_section(student_name, uni, major, target_desc, api_key, show_preview=False)


# =====================================================================  ④ 모의면접 진행 모드
RUN_GRADES = ["상", "중", "하"]


def _run_timer_html(total_left, q_elapsed):
    return f"""
<div style="font-family:Pretendard,sans-serif;display:flex;gap:28px;align-items:baseline;justify-content:center;color:#8a8f98">
  <span>전체 남은 시간 <b id="tl" style="font-size:1.7rem;color:#2E7DD7">--:--</b></span>
  <span>이 질문 <b id="tq" style="font-size:1.7rem;color:#2E7DD7">0:00</b></span>
</div>
<script>
let L={int(total_left)},Q={int(q_elapsed)};const tl=document.getElementById('tl'),tq=document.getElementById('tq');
const f=s=>{{const n=Math.abs(s);return (s<0?'-':'')+Math.floor(n/60)+':'+String(n%60).padStart(2,'0');}};
function d(){{tl.textContent=f(L);tl.style.color=L<=0?'#c62828':(L<=60?'#e65100':'#2E7DD7');
tq.textContent=f(Q);tq.style.color=Q>=90?'#c62828':(Q>=60?'#e65100':'#2E7DD7');}}
d();setInterval(()=>{{L--;Q++;d();}},1000);
</script>"""


def _interview_minutes(uni, major, default=10):
    try:
        info = get_univ_info(RT.univ_info_db, uni, major)
        mins = [int(x) for x in re.findall(r"(\d+)\s*분", str(info.get("면접시간", "")))]
        return max(mins) if mins else default
    except Exception:
        return default


def _run_move(step):
    run = st.session_state.get("run")
    if not run:
        return
    run["i"] = max(0, min(len(run["keys"]) - 1, run["i"] + step))
    run["q_start"] = time.time()


def _run_toggle_tail(key):
    run = st.session_state.get("run")
    if run:
        run["tails"][key] = not run["tails"].get(key)


def _run_finish():
    run = st.session_state.get("run")
    if run:
        run["done"] = True
        run["end"] = time.time()


def render_run_mode(student_name, uni, major, interview_type, target_desc):
    bank = current_bank(interview_type == "생기부 기반 면접")
    run = st.session_state.get("run")
    st.markdown("### 🎙️ 모의면접 진행")
    if not bank:
        st.info("진행할 질문이 없습니다. 먼저 면접 문항이나 질문지를 만들어 주세요.")
        return
    by = {e["key"]: e for e in bank}

    # ── 시작 전 ─────────────────────────────────────────────
    if not run or not any(k in by for k in run.get("keys", [])):
        picked = bank_selected_entries(bank)
        where = st.segmented_control("어디서 진행할까요?", ["📱 패드·폰으로 (면접실)", "💻 이 컴퓨터에서"],
                                     key="run_where", default="📱 패드·폰으로 (면접실)")
        if where != "💻 이 컴퓨터에서":
            render_device_run_panel(student_name, uni, major, interview_type, picked)
            if interview_type == "생기부 기반 면접" and st.button("🏦 질문 다시 고르기", key="run_pick_dev"):
                goto_section("🏦 질문 은행·질문지")
                st.rerun()
            return
        st.caption("질문을 하나씩 크게 띄우고, 대학 면접 시간으로 시간을 잽니다. 질문마다 상·중·하와 메모를 남기면 "
                   "끝낼 때 **평가표의 문항별 평가란**에 그대로 들어갑니다.")
        with st.container(border=True):
            st.markdown(f"**진행할 질문 {len(picked)}개**"
                        + (" — 🏦 질문 은행에서 고른 질문" if interview_type == "생기부 기반 면접" else ""))
            for i, e in enumerate(picked[:8], 1):
                st.caption(f"{i}. {e['q'][:70]}{'…' if len(e['q']) > 70 else ''}")
            if len(picked) > 8:
                st.caption(f"… 외 {len(picked) - 8}개")
            c1, c2 = st.columns([1, 2], vertical_alignment="bottom")
            mins = c1.number_input("면접 시간(분)", 3, 60, _interview_minutes(uni, major), key="run_minutes",
                                   help="지원 대학의 실제 면접 시간으로 기본값을 정했습니다.")
            if c2.button("▶ 모의면접 시작", type="primary", width="stretch", disabled=not picked, key="run_start"):
                st.session_state["run"] = {"keys": [e["key"] for e in picked], "i": 0, "start": time.time(),
                                           "q_start": time.time(), "total": int(mins) * 60, "grades": {}, "memos": {},
                                           "tails": {}, "done": False, "at": _kst_stamp()}
                st.rerun()
        if interview_type == "생기부 기반 면접":
            if st.button("🏦 질문 다시 고르기", key="run_pick"):
                goto_section("🏦 질문 은행·질문지")
                st.rerun()
        return

    keys = [k for k in run["keys"] if k in by]
    n = len(keys)
    run["i"] = min(run["i"], n - 1)

    # ── 끝난 뒤: 결과 정리 → 평가표 ──────────────────────────
    if run.get("done"):
        rows = run_rows(run, bank)
        spent = int(run.get("end", time.time()) - run["start"])
        cnt = {g: sum(1 for r in rows if r["grade"] == g) for g in RUN_GRADES}
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("질문", f"{sum(1 for r in rows if r['grade'])}/{n} 평가")
        m2.metric("걸린 시간", f"{spent // 60}:{spent % 60:02d}", delta=f"계획 {run['total'] // 60}분", delta_color="off")
        m3.metric("상 · 중 · 하", f"{cnt['상']} · {cnt['중']} · {cnt['하']}")
        weak = [r for r in rows if r["grade"] == "하"]
        m4.metric("다시 연습할 질문", len(weak))
        st.dataframe(pd.DataFrame([{"번호": i, "질문": r["q"], "평가": r["grade"] or "-", "메모": r["memo"]}
                                   for i, r in enumerate(rows, 1)]), hide_index=True, width="stretch")
        with st.container(border=True):
            st.markdown("**📝 평가표에 넣을 종합 의견** (비워 두면 손으로 쓰는 빈칸으로 남습니다)")
            op1 = st.text_area("① 잘한 점", key="run_op1", height=80)
            op2 = st.text_area("② 보완할 점", key="run_op2", height=80,
                               placeholder=("다시 연습할 질문: " + ", ".join(f"{rows.index(r) + 1}번" for r in weak)) if weak else "")
            op3 = st.text_area("③ 다음 연습까지 지도 방향", key="run_op3", height=70)
            e1, e2 = st.columns(2)
            ev_name = e1.text_input("평가 교사", key="run_eval_name", placeholder="예: 김기섭")
            ev_date = e2.date_input("면접 일자", value=datetime.date.today(), key="run_eval_date")
            if st.button("📝 평가표에 반영하기", type="primary", width="stretch", key="run_to_eval"):
                try:
                    st.session_state["last_run"] = {"for": student_name, "rows": rows, "ops": [op1, op2, op3],
                                                    "at": run.get("at", _kst_stamp())}
                    st.session_state.evaluation_sheet_file = create_evaluation_sheet_word(
                        st.session_state.get("last_result_text", ""), student_name, uni, major, interview_type,
                        evaluator=ev_name.strip(), interview_date=ev_date.strftime("%Y. %m. %d.") if ev_date else "",
                        include_questions=True, run_rows=rows, opinions=[op1, op2, op3])
                    log = {"at": run.get("at", _kst_stamp()), "n": n, "grades": cnt, "sec": spent,
                           "rows": [{"q": r["q"][:80], "grade": r["grade"], "memo": r["memo"][:200]} for r in rows]}
                    st.session_state["qbank_runs"] = (st.session_state.get("qbank_runs") or []) + [log]
                    save_qbank_state(st.session_state.get("current_record_id"), student_name, uni, major)
                    st.toast("평가표에 반영했습니다. 위쪽 📕 교사용 묶음에서 받으세요.", icon="📝")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ 평가표를 만들지 못했습니다: {e}")
        r1, r2 = st.columns(2)
        if r1.button("↺ 같은 질문으로 다시", width="stretch", key="run_again"):
            st.session_state["run"] = dict(run, i=0, start=time.time(), q_start=time.time(), grades={}, memos={},
                                           tails={}, done=False, at=_kst_stamp())
            st.rerun()
        if r2.button("🏁 처음 화면으로", width="stretch", key="run_reset"):
            st.session_state["run"] = None
            st.rerun()
        return

    # ── 진행 중 ─────────────────────────────────────────────
    key = keys[run["i"]]
    e = by[key]
    import streamlit.components.v1 as _components
    _components.html(_run_timer_html(run["total"] - (time.time() - run["start"]), time.time() - run["q_start"]), height=52)
    st.progress((run["i"] + 1) / n, text=f"질문 {run['i'] + 1} / {n}")
    with st.container(border=True):
        st.markdown(f"<div style='font-size:1.9rem;font-weight:700;line-height:1.5;padding:.6rem .2rem'>"
                    f"{_esc_html(e['q'])}</div>", unsafe_allow_html=True)
        if run["tails"].get(key):
            tail = e.get("tail") or ""
            if tail:
                st.markdown(f"<div style='font-size:1.45rem;font-weight:600;color:#e65100;line-height:1.5;"
                            f"padding:.2rem .2rem .4rem'>🔥 {_esc_html(tail)}</div>", unsafe_allow_html=True)
            else:
                st.caption("🔥 준비된 꼬리질문이 없습니다. 답변을 듣고 직접 이어서 물어보세요.")
        src = _qb_src_line(e)
        st.caption(f"{e['src']} {e['no']}번" + (f" · 📍 {src}" if src else ""))
    nav1, nav2, nav3, nav4 = st.columns([1, 1.3, 1, 1])
    nav1.button("◀ 이전", width="stretch", on_click=_run_move, args=(-1,), disabled=run["i"] == 0, key="run_prev")
    nav2.button("🔥 꼬리질문 " + ("숨기기" if run["tails"].get(key) else "보기"), width="stretch",
                on_click=_run_toggle_tail, args=(key,), key="run_tail")
    if run["i"] < n - 1:
        nav3.button("다음 ▶", type="primary", width="stretch", on_click=_run_move, args=(1,), key="run_next")
    else:
        nav3.button("✅ 마지막 질문", width="stretch", disabled=True, key="run_last")
    nav4.button("⏹ 면접 끝내기", width="stretch", on_click=_run_finish, key="run_end")

    g_key, m_key = f"run_g_{key}", f"run_m_{key}"
    if g_key not in st.session_state and run["grades"].get(key):
        st.session_state[g_key] = run["grades"][key]
    if m_key not in st.session_state and run["memos"].get(key):
        st.session_state[m_key] = run["memos"][key]
    g1, g2 = st.columns([1, 3], vertical_alignment="center")
    with g1:
        grade = st.segmented_control("평가", RUN_GRADES, key=g_key)
    with g2:
        memo = st.text_input("메모", key=m_key, placeholder="예: 결론이 늦게 나옴 · 실험 수치 정확히 말함")
    run["grades"][key] = grade or ""
    run["memos"][key] = memo or ""
    done = sum(1 for k in keys if run["grades"].get(k))
    st.caption(f"평가한 질문 {done}/{n} · 질문을 넘겨도 평가와 메모는 남습니다.")
    with st.expander("📕 교사용: 모범답안 · 핵심 · 근거 문단"):
        if e.get("answer"):
            st.markdown(_md_safe(e["answer"]))
        if e.get("keys"):
            st.caption("핵심: " + " · ".join(e["keys"]))
        if e.get("quote"):
            st.caption(f"📍 “{e['quote']}”")
        if e.get("avoid"):
            st.caption("✖ 피할 답변: " + e["avoid"])
        if not (e.get("answer") or e.get("quote")):
            st.caption("이 질문에는 모범답안이 없습니다.")


def _app_base_url():
    try:
        u = st.secrets.get("APP_URL")
        if u:
            return str(u).rstrip("/")
    except Exception:
        pass
    try:
        from urllib.parse import urlsplit
        p = urlsplit(str(st.context.url))
        if p.scheme and p.netloc:
            return f"{p.scheme}://{p.netloc}{p.path}".rstrip("/")
    except Exception:
        pass
    return ""


def _qr_png(url):
    try:
        import qrcode
        buf = io.BytesIO()
        qrcode.make(url).save(buf, format="PNG")
        return buf.getvalue()
    except Exception:
        return None


def _timer_html(seconds=60):
    return f"""
<div style="font-family:Pretendard,sans-serif;display:flex;gap:10px;align-items:center">
  <button id="tbtn" style="padding:6px 14px;border-radius:8px;border:1px solid #888;background:#fff;cursor:pointer">⏱️ {seconds}초 타이머 시작</button>
  <span id="tleft" style="font-size:1.4rem;font-weight:700">{seconds}</span><span>초</span>
</div>
<script>
let t=null;const b=document.getElementById('tbtn'),l=document.getElementById('tleft');
b.onclick=()=>{{if(t)clearInterval(t);let s={seconds};l.style.color='';l.textContent=s;
t=setInterval(()=>{{s--;l.textContent=s;if(s<=15)l.style.color='#e65100';if(s<=0){{clearInterval(t);l.style.color='#c62828';l.textContent='끝';}}}},1000);}};
</script>"""


def _score_bar_md(scores):
    return "  \n".join(f"{k} {'●' * v}{'○' * (5 - v)} {v}/5" for k, v in scores.items())


def _render_feedback(fb, prev_total=None, compact=False):
    if not fb:
        return
    delta = None if prev_total is None else fb["total"] - prev_total
    c1, c2 = st.columns([1, 2])
    with c1:
        st.metric("첨삭 점수", f"{fb['total']}/20", delta=(f"{delta:+d}" if delta else None))
    with c2:
        st.markdown(_score_bar_md(fb["scores"]))
    if fb.get("mismatch"):
        st.warning("⚠️ **생기부와 다른 내용이 있어요.** 면접관이 확인하면 곤란할 수 있으니 고치거나 선생님께 확인하세요.  \n"
                   + "  \n".join(f"· “{_md_safe(m)}”" for m in fb["mismatch"]))
    if fb.get("good") and not compact:
        st.markdown(f"**👍 잘한 점** {_md_safe(fb['good'])}")
    if fb.get("fixes"):
        st.markdown("**🔧 고칠 점**  \n" + "  \n".join(f"{i}. {_md_safe(x)}" for i, x in enumerate(fb["fixes"], 1)))
    if fb.get("better") and not compact:
        with st.expander("✨ 다듬은 답안 보기"):
            st.markdown(_md_safe(fb["better"]))
    if fb.get("transcript") and not compact:
        with st.expander("🗣️ 내가 말한 내용 (받아쓰기)"):
            st.markdown(_md_safe(fb["transcript"]))


# ---------------------------------------------------------------- 학생 화면
def _student_api_key():
    try:
        return st.secrets["GEMINI_API_KEY"]
    except Exception:
        return ""


def render_student_page(s_param):
    """학생이 링크로 들어왔을 때의 화면. 선생님 화면·다른 학생 기록에는 접근하지 않습니다."""
    sid = re.sub(r"[^A-Za-z0-9._-]", "", str(s_param))[:80]
    link_key = f"stu_link_{sid}"
    if link_key not in st.session_state or st.session_state.get(f"stu_reload_{sid}"):
        st.session_state.pop(f"stu_reload_{sid}", None)
        row, obj = _load_json_row(f"{STU_PREFIX}{sid}")
        st.session_state[link_key] = (dict(row) if row else None, obj)
    row, link = st.session_state[link_key]
    if not row or not link.get("record_id") or not link.get("qs"):
        st.error("링크가 올바르지 않거나, 선생님이 연습을 끝냈습니다. 선생님께 새 링크를 받아 주세요.")
        st.stop()
    name = str(_row_get(row, "student_name", ""))
    uni = str(_row_get(row, "university", ""))
    major = str(_row_get(row, "major", ""))
    record_id = link["record_id"]

    st.markdown(f"## 📱 {name} 학생 면접 연습")
    st.caption(f"{uni} · {major}")

    ok_key = f"stu_ok_{sid}"
    if not st.session_state.get(ok_key):
        locked = st.session_state.get(f"stu_lock_{sid}", 0)
        if time.time() < locked:
            st.error(f"PIN을 여러 번 틀렸습니다. {int(locked - time.time()) + 1}초 뒤에 다시 시도해 주세요.")
            st.stop()
        with st.form("stu_pin_form"):
            pin = st.text_input("선생님께 받은 PIN 4자리", type="password", max_chars=4)
            go = st.form_submit_button("시작하기", width="stretch")
        if go:
            if hmac.compare_digest(str(pin or "").strip().encode(), str(link.get("pin", "")).encode()):
                st.session_state[ok_key] = True
                st.session_state[f"stu_fail_{sid}"] = 0
                st.rerun()
            fails = st.session_state.get(f"stu_fail_{sid}", 0) + 1
            st.session_state[f"stu_fail_{sid}"] = fails
            if fails >= 5:
                st.session_state[f"stu_lock_{sid}"] = time.time() + 60
                st.session_state[f"stu_fail_{sid}"] = 0
            time.sleep(1)
            st.error("PIN이 맞지 않습니다.")
        st.stop()

    api_key = _student_api_key()
    ans_key, cmt_key = f"stu_ans_{sid}", f"stu_cmt_{sid}"
    if ans_key not in st.session_state:
        st.session_state[ans_key] = load_answers(record_id)
        st.session_state[cmt_key] = load_comments(record_id)
    ans, cmts = st.session_state[ans_key], st.session_state[cmt_key]
    limit = int(link.get("limit", STU_DAILY_LIMIT) or STU_DAILY_LIMIT)
    used = int(ans["usage"].get(_kst_today(), 0))
    qs = link["qs"]

    top1, top2 = st.columns([3, 1], vertical_alignment="center")
    with top1:
        done = sum(1 for q in qs if ans["q"].get(q["key"], {}).get("tries"))
        st.progress(done / len(qs) if qs else 0.0, text=f"답한 질문 {done}/{len(qs)} · 오늘 첨삭 {used}/{limit}회")
    with top2:
        if st.button("🔄 새로고침", width="stretch", help="선생님 코멘트와 새 질문을 다시 불러옵니다."):
            for k in (ans_key, cmt_key):
                st.session_state.pop(k, None)
            st.session_state[f"stu_reload_{sid}"] = True
            st.rerun()
    if not storage_is_shared():
        st.warning("⚠️ 이 앱은 아직 구글시트 저장이 연결되지 않아, 답안이 선생님께 전달되지 않을 수 있습니다.")
    if not api_key:
        st.error("AI 첨삭 설정이 되어 있지 않습니다. 선생님께 알려 주세요.")

    tab_practice, tab_record = st.tabs(["📝 질문 연습", "📈 내 기록"])
    with tab_practice:
        _student_practice_tab(sid, link, ans, cmts, name, uni, major, record_id, api_key, limit)
    with tab_record:
        _student_record_tab(qs, ans)


def _student_ctx(link, q, uni, major):
    return {"q": q["q"], "quote": q["quote"], "para": q["para"], "record": link.get("record", ""),
            "uni": uni, "major": major}


def _student_save(sid, record_id, name, uni, major, ans, owner):
    try:
        save_answers(record_id, name, uni, major, ans, owner=owner)
        return True
    except Exception as e:
        st.error(f"답안을 저장하지 못했습니다: {e}")
        return False


def _student_practice_tab(sid, link, ans, cmts, name, uni, major, record_id, api_key, limit):
    qs = link["qs"]
    owner = link.get("owner", "")
    flt = st.radio("보기", ["전체", "아직 안 한 질문", "다시 볼 질문"], horizontal=True, key=f"stu_flt_{sid}",
                   label_visibility="collapsed")

    def needs_review(q):
        tries = ans["q"].get(q["key"], {}).get("tries", [])
        if not tries:
            return False
        fb = tries[-1].get("fb", {})
        return fb.get("total", 20) < 14 or bool(fb.get("mismatch")) or q["key"] in cmts["c"]

    if flt == "아직 안 한 질문":
        view = [q for q in qs if not ans["q"].get(q["key"], {}).get("tries")]
    elif flt == "다시 볼 질문":
        view = [q for q in qs if needs_review(q)]
    else:
        view = list(qs)
    if not view:
        st.info("이 보기에 해당하는 질문이 없습니다.")
        return

    def label(q):
        tries = ans["q"].get(q["key"], {}).get("tries", [])
        mark = f"✅ {tries[-1].get('fb', {}).get('total', '-')}점" if tries else "⬜"
        cm = " 💬" if q["key"] in cmts["c"] else ""
        return f"{mark}{cm} {q['no']}. {q['q'][:34]}{'…' if len(q['q']) > 34 else ''}"

    keys = [q["key"] for q in view]
    sel_key = f"stu_sel_{sid}_{flt}"
    if st.session_state.get(sel_key) not in keys:
        st.session_state[sel_key] = keys[0]

    def move(step):
        i = keys.index(st.session_state[sel_key])
        st.session_state[sel_key] = keys[(i + step) % len(keys)]

    st.selectbox("질문 고르기", keys, key=sel_key, format_func=lambda k: label(next(q for q in view if q["key"] == k)))
    n1, n2 = st.columns(2)
    n1.button("◀ 이전 질문", width="stretch", on_click=move, args=(-1,), key=f"stu_prev_{sid}")
    n2.button("다음 질문 ▶", width="stretch", on_click=move, args=(1,), key=f"stu_next_{sid}")

    q = next(x for x in view if x["key"] == st.session_state[sel_key])
    item = ans["q"].setdefault(q["key"], {"q": q["q"], "tries": []})
    item["q"] = q["q"]
    tries = item["tries"]
    used = int(ans["usage"].get(_kst_today(), 0))
    over = used >= limit
    ctx = _student_ctx(link, q, uni, major)

    with st.container(border=True):
        st.markdown(f"#### {q['no']}. {_md_safe(q['q'])}")
        st.caption(f"📂 {q['area']}" + (f" · 근거 “{_md_safe(q['quote'])}”" if q.get("quote") else ""))
        if q["key"] in cmts["c"]:
            c = cmts["c"][q["key"]]
            st.info(f"👩‍🏫 **선생님 코멘트** ({c.get('at', '')})  \n{_md_safe(c.get('text', ''))}")

        mode = st.radio("답하는 방법", ["✍️ 글로 답하기", "🎙️ 말로 답하기"], horizontal=True, key=f"stu_mode_{sid}")
        if over:
            st.warning(f"오늘 첨삭을 {limit}회 모두 썼습니다. 내일 다시 할 수 있어요. 답을 다듬어 두고 내일 첨삭받아 보세요.")
        submitted, payload = False, {}
        if mode.startswith("✍️"):
            txt = st.text_area("내 답변", key=f"stu_txt_{sid}_{q['key']}", height=200,
                               placeholder="결론부터 말해 보세요. 예) 저는 ~했습니다. 그 과정에서 ~. 이를 통해 ~를 배웠습니다.")
            st.caption(f"{len(txt)}자 · 말하면 약 {max(1, round(len(txt) / 6))}초 (권장 40∼60초, 240∼360자)")
            submitted = st.button("📝 첨삭 받기", type="primary", width="stretch",
                                  disabled=over or not api_key, key=f"stu_go_{sid}_{q['key']}")
            if submitted and len(txt.strip()) < 20:
                st.warning("답변을 조금 더 써 주세요(20자 이상).")
                submitted = False
            payload = {"answer": txt.strip()}
        else:
            import streamlit.components.v1 as _components
            _components.html(_timer_html(60), height=50)
            audio = st.audio_input("녹음 버튼을 누르고 답해 보세요", key=f"stu_aud_{sid}_{q['key']}")
            data = audio.getvalue() if audio is not None else None
            sec = _wav_seconds(data) if data else None
            if sec:
                note = "적당해요" if 40 <= sec <= 70 else ("조금 짧아요" if sec < 40 else "조금 길어요")
                st.caption(f"녹음 길이 {sec}초 — {note} (권장 40∼60초)")
            submitted = st.button("📝 첨삭 받기", type="primary", width="stretch",
                                  disabled=over or not api_key, key=f"stu_go_a_{sid}_{q['key']}")
            if submitted and not data:
                st.warning("먼저 녹음 버튼을 눌러 답을 녹음해 주세요.")
                submitted = False
            payload = {"audio": data, "sec": sec}
        if submitted:
            with st.spinner("AI 선생님이 답변을 읽고 있어요... (10~30초)"):
                fb, err = run_feedback(ctx, api_key, answer=payload.get("answer", ""),
                                       audio=payload.get("audio"), sec=payload.get("sec"))
            if fb:
                tries.append({"at": _kst_stamp(), "mode": "말" if payload.get("audio") else "글",
                              "text": payload.get("answer") or fb.get("transcript", ""), "sec": payload.get("sec"),
                              "fb": fb, "follow": []})
                ans["usage"][_kst_today()] = used + 1
                _student_save(sid, record_id, name, uni, major, ans, owner)
                st.rerun()
            else:
                st.error(err)

    if not tries:
        st.caption("답을 한 번 첨삭받으면 모범답안을 볼 수 있어요. 먼저 스스로 답해 보세요!")
        return

    last = tries[-1]
    prev_total = tries[-2]["fb"].get("total") if len(tries) > 1 else None
    with st.container(border=True):
        st.markdown(f"**🧾 {len(tries)}차 첨삭 결과** · {last.get('at', '')} · {last.get('mode', '글')}로 답함"
                    + (f" · {last['sec']}초" if last.get("sec") else ""))
        _render_feedback(last.get("fb"), prev_total)
        st.caption("고칠 점을 반영해 위에서 다시 답하면 다음 차수로 쌓입니다.")

    # 꼬리질문 이어 답하기
    follows = last.setdefault("follow", [])
    for i, f in enumerate(follows, 1):
        with st.container(border=True):
            st.markdown(f"**🔥 꼬리질문 {i}.** {_md_safe(f['q'])}")
            st.markdown(f"> {_md_safe(f.get('a', ''))}")
            _render_feedback(f.get("fb"), compact=True)
    next_q = (follows[-1].get("fb", {}).get("follow") if follows else last.get("fb", {}).get("follow")) or ""
    if next_q and len(follows) < FOLLOW_MAX:
        with st.container(border=True):
            st.markdown(f"**🔥 꼬리질문 {len(follows) + 1}.** {_md_safe(next_q)}")
            fa = st.text_area("이어서 답하기", key=f"stu_fa_{sid}_{q['key']}_{len(tries)}_{len(follows)}", height=130)
            f_go = st.button("📝 꼬리질문 첨삭 받기", width="stretch", disabled=over or not api_key,
                             key=f"stu_fgo_{sid}_{q['key']}_{len(tries)}_{len(follows)}")
            if f_go and len(fa.strip()) < 10:
                st.warning("답변을 조금 더 써 주세요(10자 이상).")
                f_go = False
            if f_go:
                prev = follows[-1].get("a", "") if follows else last.get("text", "")
                with st.spinner("꼬리질문 답변을 읽고 있어요..."):
                    fb, err = run_feedback(ctx, api_key, answer=fa.strip(), follow={"q": next_q, "prev": prev})
                if fb:
                    follows.append({"q": next_q, "a": fa.strip(), "fb": fb, "at": _kst_stamp()})
                    ans["usage"][_kst_today()] = used + 1
                    _student_save(sid, record_id, name, uni, major, ans, owner)
                    st.rerun()
                else:
                    st.error(err)
    elif follows and len(follows) >= FOLLOW_MAX:
        st.success("꼬리질문 3단계까지 모두 답했어요! 실제 면접에서도 이렇게 이어서 물어봅니다.")

    with st.expander("✅ 모범답안 보기"):
        if q.get("answer"):
            st.markdown(_md_safe(q["answer"]))
        else:
            st.caption("선생님이 아직 이 질문의 모범답안을 올리지 않았어요. 위의 '다듬은 답안'을 참고하세요.")
        if q.get("tail"):
            st.markdown(f"**🔥 예상 꼬리질문** {_md_safe(q['tail'])}")
        if q.get("guide"):
            st.caption("답변 가이드: " + q["guide"])

    if len(tries) > 1:
        with st.expander(f"🕘 지난 답변 {len(tries) - 1}개"):
            for i, t in enumerate(tries[:-1], 1):
                st.markdown(f"**{i}차** · {t.get('at', '')} · {t.get('fb', {}).get('total', '-')}점")
                st.caption(t.get("text", "")[:400])


# ---------------------------------------------------------------- 선생님: 기록별 링크·현황판
def render_teacher_practice(hl, student_name, uni, major, bank=None):
    hl = hl or {}
    record_id = st.session_state.get("current_record_id")
    st.markdown("### 📱 학생 연습 — 휴대폰·패드로 답변 받기")
    if not record_id:
        st.caption("기록을 먼저 저장하면 학생 링크를 만들 수 있습니다.")
        return
    if not storage_is_shared():
        st.warning("⚠️ 구글시트 저장이 설정되지 않아, 학생 기기에서 쓴 답안이 선생님 화면과 공유되지 않습니다. "
                   "구글시트(Apps Script) 연동 후 사용해 주세요.")
    cache = st.session_state.setdefault("tp_cache", {})
    ck = _bundle_key(record_id)
    reload = st.session_state.pop("tp_reload", False)
    if ck not in cache or reload:
        links = practice_link_rows(record_id)
        link_row, link = (None, {})
        if links:
            link_row, link = _load_json_row(_row_get(links[0], "id", ""))
        cache[ck] = {"link_id": _row_get(links[0], "id", "") if links else None, "link": link,
                     "ans": load_answers(record_id), "cmt": load_comments(record_id)}
    data = cache[ck]
    link, link_id = data["link"], data["link_id"]

    picked = bank_selected_entries(bank) if bank else []
    hl_n = sum(len(c["qs"]) for c in hl.get("items", []))
    opts = []
    if picked:
        opts.append(f"🏦 질문 은행에서 고른 질문 ({len(picked)}개)")
    if hl_n:
        opts.append(f"🖍️ 형광펜 질문 전체 ({hl_n}개)")
    if not opts:
        st.caption("학생에게 보낼 질문이 없습니다. 🏦 질문 은행에서 질문을 고르거나 형광펜 질문 카드를 만들어 주세요.")
        return
    pref = st.session_state.pop("tp_src_pref", None)
    if pref == "bank" and picked:
        st.session_state["tp_src"] = opts[0]
    if st.session_state.get("tp_src") not in opts:
        st.session_state["tp_src"] = opts[0]
    src_choice = st.radio("학생에게 보낼 질문", opts, horizontal=True, key="tp_src",
                          help="링크를 만들거나 '🔄 최신 질문 보내기'를 누를 때 이 질문이 학생 화면으로 갑니다.")

    def snapshot():
        rec = st.session_state.get("loaded_student_record_text", "")
        if src_choice.startswith("🏦"):
            return practice_snapshot_from_bank(picked, rec, student_name)
        return practice_snapshot(hl, st.session_state.get("last_result_text", ""), rec, student_name)

    with st.container(border=True):
        if not link_id:
            st.markdown("학생이 휴대폰이나 패드로 자기 질문에 답하고 AI 첨삭을 받을 수 있는 **전용 링크**를 만듭니다. "
                        "학생 화면에는 질문·근거 문장·모범답안만 가고, 학생 이름은 가려서 AI로 보냅니다.")
            lim = st.number_input("학생 하루 첨삭 횟수", 5, 200, STU_DAILY_LIMIT, step=5, key="tp_limit_new")
            if st.button("📱 학생용 링크 만들기", type="primary", key="tp_make"):
                token = _pysecrets.token_urlsafe(9).replace("-", "x").replace("_", "y")
                pin = f"{_pysecrets.randbelow(10000):04d}"
                try:
                    save_student_link(token, record_id, student_name, uni, major,
                                      {"pin": pin, "limit": int(lim), "owner": current_teacher() or "",
                                       "created": _kst_stamp()}, snapshot())
                    st.session_state["tp_reload"] = True
                    st.rerun()
                except Exception as e:
                    st.error(f"링크를 만들지 못했습니다: {e}")
        else:
            s_param = link_id[len(STU_PREFIX):]
            base = _app_base_url()
            url = f"{base}/?s={s_param}" if base else f"(앱 주소)/?s={s_param}"
            c1, c2 = st.columns([3, 1])
            with c1:
                st.markdown(f"**학생 링크** · PIN **{link.get('pin', '????')}** · 하루 {link.get('limit', STU_DAILY_LIMIT)}회")
                st.code(url, language=None)
                if not base:
                    st.caption("앱 주소를 알아내지 못했습니다. Secrets에 APP_URL = \"https://...streamlit.app\"을 넣으면 전체 주소가 나옵니다.")
                st.caption(f"학생에게 보낸 질문 {len(link.get('qs', []))}개 · 마지막 보냄 {link.get('synced', '')}")
            with c2:
                png = _qr_png(url) if base else None
                if png:
                    st.image(png, caption="QR코드", width=140)
            b1, b2, b3, b4 = st.columns(4)
            settings = {k: link.get(k) for k in ("pin", "limit", "owner", "created")}
            token = s_param.split(".", 1)[1] if "." in s_param else s_param
            if b1.button("🔄 최신 질문 보내기", width="stretch", key="tp_sync",
                         help="질문을 더 만들었거나 모범답안을 만든 뒤 누르면 학생 화면에 반영됩니다."):
                try:
                    save_student_link(token, record_id, student_name, uni, major, settings, snapshot())
                    st.session_state["tp_reload"] = True
                    st.rerun()
                except Exception as e:
                    st.error(f"보내지 못했습니다: {e}")
            if b2.button("🔢 PIN 바꾸기", width="stretch", key="tp_pin"):
                settings["pin"] = f"{_pysecrets.randbelow(10000):04d}"
                save_student_link(token, record_id, student_name, uni, major, settings,
                                  {"qs": link.get("qs", []), "record": link.get("record", "")})
                st.session_state["tp_reload"] = True
                st.rerun()
            new_lim = b3.number_input("하루 횟수", 5, 200, int(link.get("limit", STU_DAILY_LIMIT)), step=5,
                                      key="tp_limit", label_visibility="collapsed")
            if new_lim != int(link.get("limit", STU_DAILY_LIMIT)):
                settings["limit"] = int(new_lim)
                save_student_link(token, record_id, student_name, uni, major, settings,
                                  {"qs": link.get("qs", []), "record": link.get("record", "")})
                st.session_state["tp_reload"] = True
                st.rerun()
            if b4.button("⛔ 링크 끊기", width="stretch", key="tp_cut",
                         help="학생이 더 이상 들어올 수 없습니다. 답안 기록은 남고, 새 링크를 만들면 이어집니다."):
                _delete_record_raw(link_id)
                st.session_state["tp_reload"] = True
                st.rerun()

    render_practice_board(hl, record_id, student_name, uni, major, data)


def render_practice_board(hl, record_id, student_name, uni, major, data):
    ans, cmt = data["ans"], data["cmt"]
    head1, head2 = st.columns([3, 1], vertical_alignment="center")
    with head1:
        st.markdown(f"**📋 학생 답안 현황판** · 마지막 활동 {ans.get('updated', '없음')}")
    with head2:
        if st.button("🔄 답안 새로고침", width="stretch", key="tp_refresh"):
            st.session_state["tp_reload"] = True
            st.rerun()
    # 학생에게 실제로 보낸 질문 순서·번호를 기준으로 봅니다(형광펜·은행 어느 쪽에서 보냈든).
    qmap = {q["key"]: q for q in (data.get("link") or {}).get("qs", [])} \
        or {q["key"]: q for c in (hl or {}).get("items", []) for q in c["qs"]}
    answered = [k for k, v in ans["q"].items() if v.get("tries")]
    if not answered:
        st.caption("아직 학생이 첨삭받은 답변이 없습니다.")
        return
    lasts = [ans["q"][k]["tries"][-1].get("fb", {}) for k in answered]
    mism = [k for k in answered if ans["q"][k]["tries"][-1].get("fb", {}).get("mismatch")]
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("답한 질문", f"{len(answered)}/{len(qmap) or len(answered)}")
    m2.metric("평균 점수(최근)", f"{round(sum(f.get('total', 0) for f in lasts) / len(lasts), 1)}/20")
    m3.metric("첨삭 횟수", sum(len(ans["q"][k]["tries"]) for k in answered))
    m4.metric("⚠️ 생기부와 다름", len(mism))
    crit = {k: round(sum(f.get("scores", {}).get(k, 0) for f in lasts) / len(lasts), 1) for k in FB_CRITERIA}
    weak = min(crit, key=crit.get)
    st.caption("기준별 평균: " + " · ".join(f"{k} {v}" for k, v in crit.items()) + f" — 가장 약한 기준: **{weak}**")

    order = [k for k in qmap if k in answered] + [k for k in answered if k not in qmap]
    for k in order:
        item = ans["q"][k]
        tries = item["tries"]
        fb = tries[-1].get("fb", {})
        no = qmap[k]["no"] if k in qmap else "-"
        flag = " ⚠️" if fb.get("mismatch") else ""
        cm = " 💬" if k in cmt["c"] else ""
        with st.expander(f"{no}. {item.get('q', '')[:40]} — {len(tries)}회 · 최근 {fb.get('total', '-')}/20{flag}{cm}"):
            t = tries[-1]
            st.markdown(f"**최근 답변** ({t.get('at', '')}, {t.get('mode', '글')}"
                        + (f", {t['sec']}초" if t.get("sec") else "") + ")")
            st.markdown(f"> {_md_safe(t.get('text', ''))}")
            _render_feedback(fb, tries[-2]["fb"].get("total") if len(tries) > 1 else None, compact=True)
            for i, f in enumerate(t.get("follow", []), 1):
                st.markdown(f"**꼬리질문 {i}.** {_md_safe(f['q'])}  \n> {_md_safe(f.get('a', ''))}  \n"
                            f":gray[{f.get('fb', {}).get('total', '-')}/20]")
            if len(tries) > 1:
                st.caption("점수 변화: " + " → ".join(str(x.get("fb", {}).get("total", "-")) for x in tries))
            cur = cmt["c"].get(k, {}).get("text", "")
            txt = st.text_area("선생님 코멘트 (학생 화면 맨 위에 보입니다)", value=cur, key=f"tp_cmt_{k}", height=90)
            if st.button("💬 코멘트 저장", key=f"tp_cmt_save_{k}"):
                if txt.strip():
                    cmt["c"][k] = {"text": txt.strip(), "at": _kst_stamp()}
                else:
                    cmt["c"].pop(k, None)
                try:
                    save_comments(record_id, student_name, uni, major, cmt)
                    st.success("저장했습니다. 학생이 '새로고침'하면 보입니다.")
                except Exception as e:
                    st.error(f"저장하지 못했습니다: {e}")


def render_class_practice_stats():
    with st.expander("📊 우리 반 학생 연습 현황 (휴대폰 답안 첨삭)", expanded=False):
        if not storage_is_shared():
            st.caption("구글시트 저장이 연결되어야 학생 기기의 답안을 모아 볼 수 있습니다.")
        if st.button("📊 학생 연습 기록 불러오기", key="class_stats_load"):
            with st.spinner("학생들의 답안 기록을 모으는 중입니다..."):
                st.session_state["class_stats"] = load_class_practice()
        data = st.session_state.get("class_stats")
        if data is None:
            st.caption("버튼을 누르면 학생들이 휴대폰으로 연습한 기록을 모아 보여 줍니다.")
            return
        rows, crits = [], []
        for d in data:
            s = _practice_stats(d["ans"])
            if not s:
                continue
            crits.append(s["crit"])
            rows.append({"학생": d["name"], "지원": f"{d['uni']} {d['major']}", "답한 질문": s["answered"],
                         "첨삭": s["tries"], "꼬리질문": s["follow"], "처음 평균": s["first_total"],
                         "최근 평균": s["last_total"], "향상": round(s["last_total"] - s["first_total"], 1),
                         "가장 약한 기준": min(s["crit"], key=s["crit"].get), "⚠️ 생기부와 다름": s["mismatch"],
                         "마지막 활동": s["updated"]})
        if not rows:
            st.info("아직 학생들이 첨삭받은 기록이 없습니다.")
            return
        cls = {k: round(sum(c[k] for c in crits) / len(crits), 2) for k in FB_CRITERIA}
        weak = min(cls, key=cls.get)
        m1, m2, m3 = st.columns(3)
        m1.metric("연습한 학생", len(rows))
        m2.metric("반 평균 (최근)", f"{round(sum(r['최근 평균'] for r in rows) / len(rows), 1)}/20")
        m3.metric("평균 향상", f"{round(sum(r['향상'] for r in rows) / len(rows), 1):+}점")
        st.markdown("**기준별 반 평균 (5점 만점)**")
        st.bar_chart(pd.DataFrame({"반 평균": [cls[k] for k in FB_CRITERIA]}, index=FB_CRITERIA), height=220, horizontal=True)
        st.caption(f"우리 반이 가장 약한 기준은 **{weak}**({FB_CRITERIA_HELP[weak]})입니다. 다음 수업 주제로 삼아 보세요.")
        st.dataframe(pd.DataFrame(rows).sort_values("최근 평균"), hide_index=True, width="stretch")


# ---------------------------------------------------------------- 📱 패드·폰 면접관 모드
# 교무실 컴퓨터에서 질문을 고르고 '면접관 링크'(QR·PIN)를 만들면, 면접실에서 패드나 폰으로
# 질문을 하나씩 띄우고 상·중·하·메모를 남깁니다. 결과는 저장소의 IVR- 줄에 쌓이고,
# 컴퓨터의 '🎙️ 모의면접 진행 → 📱 패드·폰' 화면에서 불러와 평가표를 만듭니다.
MOBILE_CSS = """
<style>
[data-testid="stSidebar"], [data-testid="stSidebarCollapsedControl"], header[data-testid="stHeader"],
[data-testid="stToolbar"], footer, .hero-banner {display:none !important;}
.block-container {padding: .8rem .7rem 5rem !important; max-width: 760px !important;}
.stButton button, .stDownloadButton button, div[data-testid="stFormSubmitButton"] button {
  min-height: 3.3rem; font-size: 1.12rem; font-weight: 600; border-radius: 12px;}
div[data-testid="stElementContainer"]:has(div[data-testid="stButtonGroup"]) {width: 100% !important;}
div[data-testid="stButtonGroup"] {width: 100% !important; display: flex !important;}
div[data-testid="stButtonGroup"] div[role="radiogroup"] {width: 100% !important; max-width: none !important;
  flex: 1 1 100% !important; display: flex !important;}
div[data-testid="stButtonGroup"] button[role="radio"] {flex: 1 1 0 !important; min-height: 3.7rem; font-weight: 700;}
div[data-testid="stButtonGroup"] button[role="radio"] * {font-size: 1.4rem !important;}
div[data-testid="stButtonGroup"] button[role="radio"]:nth-child(1)[aria-checked="true"] {background: #2E7D32 !important; border-color: #2E7D32 !important;}
div[data-testid="stButtonGroup"] button[role="radio"]:nth-child(2)[aria-checked="true"] {background: #F9A825 !important; border-color: #F9A825 !important;}
div[data-testid="stButtonGroup"] button[role="radio"]:nth-child(3)[aria-checked="true"] {background: #C62828 !important; border-color: #C62828 !important;}
div[data-testid="stButtonGroup"] button[role="radio"][aria-checked="true"] * {color: #fff !important;}
div[data-testid="stHorizontalBlock"] {flex-wrap: nowrap !important; gap: .45rem !important;}
div[data-testid="stColumn"] {min-width: 0 !important; width: auto !important; flex: 1 1 0 !important;}
div.stButton > button[data-testid="stBaseButton-secondary"] {background: transparent !important;
  color: inherit !important; border: 1.5px solid #cbd5e1 !important;}  /* 주된 버튼만 진하게 */
input, textarea {font-size: 16px !important;}  /* 아이폰이 입력칸을 확대하지 않게 */
</style>
"""


def interviewer_snapshot(entries):
    """면접관 화면에 보낼 질문(교사용이라 모범답안·근거까지)."""
    return [{"key": e["key"], "no": i, "q": e["q"], "tail": e.get("tail", ""), "answer": e.get("answer", ""),
             "keys": e.get("keys") or [], "quote": e.get("quote", ""), "area": _qb_src_line(e), "src": e["src"]}
            for i, e in enumerate(entries, 1)]


def device_runs_cached(record_id, reload=False):
    """패드·폰 진행 결과(화면을 그릴 때마다 저장소를 읽지 않도록 세션에 보관)."""
    if not record_id:
        return []
    cache = st.session_state.setdefault("iv_runs_cache", {})
    ck = _bundle_key(record_id)
    if reload or ck not in cache:
        try:
            cache[ck] = load_device_runs(record_id)
        except Exception:
            cache[ck] = cache.get(ck, [])
    return cache[ck]


def _device_run_record(run, qs, status):
    """진행 상태 → 저장할 결과 한 건."""
    rows = [{"q": q["q"], "grade": run["grades"].get(q["key"], ""), "memo": run["memos"].get(q["key"], "")}
            for q in qs if q["key"] in run["keys"]]
    cnt = {g: sum(1 for r in rows if r["grade"] == g) for g in RUN_GRADES}
    return {"id": run["id"], "link": run.get("link", ""), "at": run["at"], "status": status,
            "i": run["i"], "n": len(rows), "total": run["total"], "sec": int(time.time() - run["start"]),
            "grades": dict(run["grades"]), "memos": dict(run["memos"]), "counts": cnt, "rows": rows,
            "ops": run.get("ops") or ["", "", ""], "evaluator": run.get("evaluator", ""), "device": "패드·폰"}


def render_device_run_panel(student_name, uni, major, interview_type, picked):
    """컴퓨터 화면: 면접관 링크(QR·PIN) 만들기와 패드·폰 결과 받아 평가표 만들기."""
    record_id = st.session_state.get("current_record_id")
    st.caption("교무실 컴퓨터에서 질문을 고른 뒤 **면접관 링크**를 만들고, 면접실에서 패드나 폰으로 QR을 찍어 진행하세요. "
               "끝나면 결과가 이 화면으로 돌아와 평가표를 만들 수 있습니다.")
    if not record_id:
        st.info("먼저 위쪽 **💾 저장**을 눌러 이 학생 기록을 저장해 주세요. 저장된 기록에만 링크를 만들 수 있습니다.")
        return
    if not storage_is_shared():
        st.warning("⚠️ 구글시트 저장이 설정되지 않아, 패드·폰에서 진행한 결과가 이 컴퓨터로 돌아오지 않습니다. "
                   "구글시트(Apps Script) 연동 후 사용해 주세요.")
    cache = st.session_state.setdefault("iv_link_cache", {})
    ck = _bundle_key(record_id)
    if st.session_state.pop("iv_reload", False) or ck not in cache:
        rows = interviewer_link_rows(record_id)
        link_id, link = (None, {})
        if rows:
            link_id = _row_get(rows[0], "id", "")
            _, link = _load_json_row(link_id)
        cache[ck] = (link_id, link)
    link_id, link = cache[ck]

    with st.container(border=True):
        if not link_id:
            st.markdown(f"**📱 면접관 링크 만들기** — 고른 질문 {len(picked)}개")
            c1, c2 = st.columns([1, 2], vertical_alignment="bottom")
            mins = c1.number_input("면접 시간(분)", 3, 60, _interview_minutes(uni, major), key="iv_minutes")
            if c2.button("📱 면접관 링크 만들기", type="primary", width="stretch", disabled=not picked, key="iv_make"):
                token = _pysecrets.token_urlsafe(9).replace("-", "x").replace("_", "y")
                try:
                    save_interviewer_link(token, record_id, student_name, uni, major, {
                        "pin": f"{_pysecrets.randbelow(10000):04d}", "owner": current_teacher() or "",
                        "created": _kst_stamp(), "minutes": int(mins), "type": interview_type,
                        "qs": interviewer_snapshot(picked)})
                    st.session_state["iv_reload"] = True
                    st.rerun()
                except Exception as e:
                    st.error(f"링크를 만들지 못했습니다: {e}")
        else:
            s_param = link_id[len(IVW_PREFIX):]
            base = _app_base_url()
            url = f"{base}/?iv={s_param}" if base else f"(앱 주소)/?iv={s_param}"
            q1, q2 = st.columns([1, 2], vertical_alignment="center")
            with q1:
                png = _qr_png(url) if base else None
                if png:
                    st.image(png, width=190)
            with q2:
                st.markdown(f"#### PIN **{link.get('pin', '????')}**")
                st.caption(f"질문 {len(link.get('qs', []))}개 · 면접 {link.get('minutes', 10)}분 · 보낸 시각 {link.get('synced', '')}")
                st.code(url, language=None)
                if not base:
                    st.caption("앱 주소를 알아내지 못했습니다. Secrets에 APP_URL = \"https://...streamlit.app\"을 넣으면 QR이 나옵니다.")
            b1, b2, b3 = st.columns(3)
            token = s_param.split(".", 1)[1] if "." in s_param else s_param
            keep = {k: link.get(k) for k in ("pin", "owner", "created", "minutes", "type")}
            if b1.button("🔄 고른 질문 다시 보내기", width="stretch", disabled=not picked, key="iv_sync",
                         help="질문 은행에서 질문을 바꿨다면 눌러 주세요. 진행 중인 패드·폰은 '처음 화면'으로 돌아가면 반영됩니다."):
                save_interviewer_link(token, record_id, student_name, uni, major, dict(keep, qs=interviewer_snapshot(picked)))
                st.session_state["iv_reload"] = True
                st.rerun()
            if b2.button("🔢 PIN 바꾸기", width="stretch", key="iv_pin"):
                keep["pin"] = f"{_pysecrets.randbelow(10000):04d}"
                save_interviewer_link(token, record_id, student_name, uni, major, dict(keep, qs=link.get("qs", [])))
                st.session_state["iv_reload"] = True
                st.rerun()
            if b3.button("⛔ 링크 끊기", width="stretch", key="iv_cut", help="결과 기록은 남습니다."):
                _delete_record_raw(link_id)
                st.session_state["iv_reload"] = True
                st.rerun()

    h1, h2 = st.columns([3, 1], vertical_alignment="center")
    h1.markdown("**📥 패드·폰에서 진행한 결과**")
    reload = h2.button("🔄 새로고침", width="stretch", key="iv_runs_reload")
    runs = device_runs_cached(record_id, reload=reload)
    if not runs:
        st.caption("아직 받은 결과가 없습니다. 면접이 끝나고 패드·폰에서 '저장하고 마치기'를 누른 뒤 새로고침해 주세요.")
        return
    for r in reversed(runs):
        cnt = r.get("counts") or {}
        graded = sum(1 for x in r.get("rows", []) if x.get("grade"))
        done = r.get("status") == "완료"
        with st.container(border=True):
            st.markdown(f"**{r.get('at', '')}** · {'✅ 완료' if done else '⏳ 진행 중'} · 평가 {graded}/{r.get('n', 0)} · "
                        f"상 {cnt.get('상', 0)} · 중 {cnt.get('중', 0)} · 하 {cnt.get('하', 0)}"
                        + (f" · 면접관 {r['evaluator']}" if r.get("evaluator") else ""))
            with st.expander("질문별 평가 · 메모 보기"):
                st.dataframe(pd.DataFrame([{"번호": i, "질문": x["q"], "평가": x.get("grade") or "-", "메모": x.get("memo", "")}
                                           for i, x in enumerate(r.get("rows", []), 1)]), hide_index=True, width="stretch")
                ops = [o for o in (r.get("ops") or []) if o]
                if ops:
                    st.caption("종합 의견: " + " / ".join(ops))
            if st.button("📝 이 결과로 평가표 만들기", key=f"iv_eval_{r.get('id')}", width="stretch", type="primary" if done else "secondary"):
                try:
                    st.session_state["last_run"] = {"for": student_name, "rows": r.get("rows", []), "ops": r.get("ops"),
                                                    "at": r.get("at", "")}
                    st.session_state.evaluation_sheet_file = create_evaluation_sheet_word(
                        st.session_state.get("last_result_text", ""), student_name, uni, major, interview_type,
                        evaluator=r.get("evaluator", ""), interview_date=str(r.get("at", ""))[:10].replace("-", ". ") + ".",
                        include_questions=True, run_rows=r.get("rows", []), opinions=r.get("ops"))
                    st.toast("평가표를 만들었습니다. 위쪽 📕 교사용 묶음에서 받으세요.", icon="📝")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ 평가표를 만들지 못했습니다: {e}")


def _iv_save(sid, run, qs, status, row_info):
    """패드·폰 진행 상태를 저장소에 남깁니다(실패해도 진행은 계속)."""
    name, uni, major = row_info
    try:
        save_device_run(run["record_id"], name, uni, major, _device_run_record(run, qs, status))
        run["saved_at"] = _kst_stamp()
        run.pop("save_error", None)
    except Exception as e:
        run["save_error"] = str(e)[:120]


def _iv_move(sid, step, qs, row_info):
    run = st.session_state.get(f"iv_run_{sid}")
    if not run:
        return
    run["i"] = max(0, min(len(run["keys"]) - 1, run["i"] + step))
    run["q_start"] = time.time()
    _iv_save(sid, run, qs, "진행 중", row_info)     # 폰 화면이 꺼져도 이어서 할 수 있게 넘길 때마다 저장


def _iv_tail(sid, key):
    run = st.session_state.get(f"iv_run_{sid}")
    if run:
        run["tails"][key] = not run["tails"].get(key)


def _iv_finish(sid):
    run = st.session_state.get(f"iv_run_{sid}")
    if run:
        run["done"] = True


def render_interviewer_page(param):
    """면접실에서 패드·폰으로 여는 면접관 화면(교사 로그인 없이 링크 + PIN)."""
    st.markdown(MOBILE_CSS, unsafe_allow_html=True)
    sid = re.sub(r"[^A-Za-z0-9._-]", "", str(param))[:80]
    lk = f"iv_link_{sid}"
    if lk not in st.session_state or st.session_state.pop(f"iv_relink_{sid}", False):
        row, obj = _load_json_row(f"{IVW_PREFIX}{sid}")
        st.session_state[lk] = (dict(row) if row else None, obj)
    row, link = st.session_state[lk]
    if not row or not link.get("qs") or not link.get("record_id"):
        st.error("링크가 올바르지 않거나 끊긴 링크입니다. 교무실 컴퓨터에서 면접관 링크를 다시 만들어 주세요.")
        st.stop()
    name = str(_row_get(row, "student_name", ""))
    uni, major = str(_row_get(row, "university", "")), str(_row_get(row, "major", ""))
    row_info = (name, uni, major)
    qs = link["qs"]
    by = {q["key"]: q for q in qs}

    ok = f"iv_ok_{sid}"
    if not st.session_state.get(ok):
        st.markdown(f"### 🎙️ {_md_safe(name)} 모의면접")
        st.caption(f"{uni} · {major} · 질문 {len(qs)}개 · {link.get('minutes', 10)}분")
        locked = st.session_state.get(f"iv_lock_{sid}", 0)
        if time.time() < locked:
            st.error(f"PIN을 여러 번 틀렸습니다. {int(locked - time.time()) + 1}초 뒤에 다시 시도해 주세요.")
            st.stop()
        with st.form("iv_pin_form"):
            pin = st.text_input("면접관 PIN 4자리 (교무실 컴퓨터 화면에 있음)", type="password", max_chars=4)
            go = st.form_submit_button("시작하기", width="stretch")
        if go:
            if hmac.compare_digest(str(pin or "").strip().encode(), str(link.get("pin", "")).encode()):
                st.session_state[ok] = True
                st.rerun()
            fails = st.session_state.get(f"iv_fail_{sid}", 0) + 1
            st.session_state[f"iv_fail_{sid}"] = fails
            if fails >= 5:
                st.session_state[f"iv_lock_{sid}"] = time.time() + 60
                st.session_state[f"iv_fail_{sid}"] = 0
            time.sleep(1)
            st.error("PIN이 맞지 않습니다.")
        st.stop()

    rk = f"iv_run_{sid}"
    run = st.session_state.get(rk)

    # ── 시작 화면 ───────────────────────────────────────────
    if not run:
        st.markdown(f"### 🎙️ {_md_safe(name)} 모의면접")
        st.caption(f"{uni} · {major} · 질문 {len(qs)}개 · {link.get('minutes', 10)}분")
        try:
            prev = [r for r in load_device_runs(link["record_id"]) if r.get("link") == sid and r.get("status") != "완료"]
        except Exception:
            prev = []
        ev = st.text_input("면접관 이름 (평가표에 들어감, 비워도 됨)", key=f"iv_ev_{sid}")
        mins = int(link.get("minutes", 10) or 10)

        def new_run(base=None):
            r = {"id": (base or {}).get("id") or uuid.uuid4().hex[:8], "link": sid, "record_id": link["record_id"],
                 "keys": [q["key"] for q in qs], "i": 0, "start": time.time(), "q_start": time.time(),
                 "total": mins * 60, "grades": {}, "memos": {}, "tails": {}, "done": False,
                 "at": (base or {}).get("at") or _kst_stamp(), "evaluator": ev.strip(), "ops": ["", "", ""]}
            if base:
                r.update(i=min(int(base.get("i", 0)), len(qs) - 1), grades=dict(base.get("grades") or {}),
                         memos=dict(base.get("memos") or {}), start=time.time() - int(base.get("sec", 0)),
                         evaluator=ev.strip() or base.get("evaluator", ""))
            return r

        if prev:
            p = prev[-1]
            st.info(f"⏳ {p.get('at', '')}에 하던 면접이 있습니다 — {int(p.get('i', 0)) + 1}번 질문까지 진행")
            if st.button("▶ 이어서 하기", type="primary", width="stretch", key=f"iv_resume_{sid}"):
                st.session_state[rk] = new_run(p)
                st.rerun()
        if st.button("▶ 면접 시작" if not prev else "새로 시작", type="secondary" if prev else "primary",
                     width="stretch", key=f"iv_start_{sid}"):
            st.session_state[rk] = new_run()
            _iv_save(sid, st.session_state[rk], qs, "진행 중", row_info)
            st.rerun()
        with st.expander(f"질문 {len(qs)}개 미리 보기"):
            for q in qs:
                st.caption(f"{q['no']}. {q['q']}")
        st.caption("💡 면접 중에는 화면 자동 꺼짐 시간을 길게 해 두세요. 화면이 꺼져도 질문을 넘길 때마다 저장되어 '이어서 하기'로 돌아올 수 있습니다.")
        if st.button("🔄 질문 다시 받기", width="stretch", key=f"iv_relink_btn_{sid}",
                     help="교무실 컴퓨터에서 질문을 바꿔 보냈다면 눌러 주세요."):
            st.session_state[f"iv_relink_{sid}"] = True
            st.rerun()
        return

    keys = [k for k in run["keys"] if k in by]
    n = len(keys)
    run["i"] = min(run["i"], n - 1)

    # ── 끝내기 화면 ─────────────────────────────────────────
    if run.get("done"):
        rec = _device_run_record(run, qs, "완료")
        cnt = rec["counts"]
        st.markdown(f"### ✅ {_md_safe(name)} 면접 끝")
        c1, c2, c3 = st.columns(3)
        c1.metric("평가", f"{sum(1 for x in rec['rows'] if x['grade'])}/{n}")
        c2.metric("걸린 시간", f"{rec['sec'] // 60}:{rec['sec'] % 60:02d}")
        c3.metric("상·중·하", f"{cnt['상']}·{cnt['중']}·{cnt['하']}")
        for i, x in enumerate(rec["rows"], 1):
            st.caption(f"{i}. [{x['grade'] or '-'}] {x['q'][:50]}" + (f" — {x['memo']}" if x["memo"] else ""))
        ops = run.setdefault("ops", ["", "", ""])
        ops[0] = st.text_area("① 잘한 점", value=ops[0], key=f"iv_op1_{sid}", height=90)
        ops[1] = st.text_area("② 보완할 점", value=ops[1], key=f"iv_op2_{sid}", height=90)
        ops[2] = st.text_area("③ 다음 연습까지 지도 방향", value=ops[2], key=f"iv_op3_{sid}", height=80)
        if st.button("💾 저장하고 마치기", type="primary", width="stretch", key=f"iv_save_{sid}"):
            _iv_save(sid, run, qs, "완료", row_info)
            if run.get("save_error"):
                st.error(f"저장하지 못했습니다: {run['save_error']} — 다시 눌러 주세요.")
            else:
                st.session_state[rk] = None
                st.success("저장했습니다. 교무실 컴퓨터의 🎙️ 모의면접 진행 → 📱 패드·폰 화면에서 '새로고침' 후 평가표를 만드세요.")
                st.balloons()
        if st.button("◀ 질문으로 돌아가기", width="stretch", key=f"iv_back_{sid}"):
            run["done"] = False
            st.rerun()
        return

    # ── 진행 중 ────────────────────────────────────────────
    key = keys[run["i"]]
    q = by[key]
    import streamlit.components.v1 as _components
    _components.html(_run_timer_html(run["total"] - (time.time() - run["start"]), time.time() - run["q_start"]), height=46)
    st.progress((run["i"] + 1) / n, text=f"{_md_safe(name)} · 질문 {run['i'] + 1} / {n}")
    with st.container(border=True):
        st.markdown(f"<div style='font-size:1.55rem;font-weight:700;line-height:1.5;padding:.3rem .1rem'>"
                    f"{_esc_html(q['q'])}</div>", unsafe_allow_html=True)
        if run["tails"].get(key):
            if q.get("tail"):
                st.markdown(f"<div style='font-size:1.3rem;font-weight:600;color:#e65100;line-height:1.5;padding:.2rem .1rem'>"
                            f"🔥 {_esc_html(q['tail'])}</div>", unsafe_allow_html=True)
            else:
                st.caption("🔥 준비된 꼬리질문이 없습니다. 답변을 듣고 직접 이어서 물어보세요.")
        if q.get("area"):
            st.caption(f"📍 {q['area']}")
    g_key, m_key = f"iv_g_{sid}_{key}", f"iv_m_{sid}_{key}"
    if g_key not in st.session_state and run["grades"].get(key):
        st.session_state[g_key] = run["grades"][key]
    if m_key not in st.session_state and run["memos"].get(key):
        st.session_state[m_key] = run["memos"][key]
    grade = st.segmented_control("평가", RUN_GRADES, key=g_key, label_visibility="collapsed")
    memo = st.text_input("메모", key=m_key, placeholder="메모 (키보드 🎤로 말해서 입력해도 됩니다)",
                         label_visibility="collapsed")
    run["grades"][key] = grade or ""
    run["memos"][key] = memo or ""
    n1, n2, n3 = st.columns([1, 1.15, 1])
    n1.button("◀", width="stretch", on_click=_iv_move, args=(sid, -1, qs, row_info), disabled=run["i"] == 0,
              key=f"iv_prev_{sid}", help="이전 질문")
    n2.button("🔥 꼬리", width="stretch", on_click=_iv_tail, args=(sid, key), key=f"iv_tail_{sid}",
              help="꼬리질문 보기/숨기기")
    if run["i"] < n - 1:
        n3.button("다음 ▶", type="primary", width="stretch", on_click=_iv_move, args=(sid, 1, qs, row_info),
                  key=f"iv_next_{sid}")
    else:
        n3.button("✅ 끝", type="primary", width="stretch", on_click=_iv_finish, args=(sid,), key=f"iv_last_{sid}")
    done = sum(1 for k in keys if run["grades"].get(k))
    st.caption(f"평가 {done}/{n}" + (f" · 저장 {run['saved_at'][11:]}" if run.get("saved_at") else "")
               + (" · ⚠️ 저장 실패(다음에 다시 시도)" if run.get("save_error") else ""))
    with st.expander("📕 모범답안 · 핵심 · 근거"):
        if q.get("answer"):
            st.markdown(_md_safe(q["answer"]))
        if q.get("keys"):
            st.caption("핵심: " + " · ".join(q["keys"]))
        if q.get("quote"):
            st.caption(f"📍 “{q['quote']}”")
        if not (q.get("answer") or q.get("quote")):
            st.caption("이 질문에는 모범답안이 없습니다.")
    with st.popover("⏹ 면접 끝내기", width="stretch"):
        st.caption("평가를 마치고 결과 화면으로 갑니다.")
        st.button("끝내기", type="primary", width="stretch", on_click=_iv_finish, args=(sid,), key=f"iv_end_{sid}")


__all__ = ["MOBILE_CSS", "interviewer_snapshot", "device_runs_cached", "_device_run_record", "render_device_run_panel", "_iv_save", "_iv_move", "_iv_tail", "_iv_finish", "render_interviewer_page", "render_interview_profile", "logout", "require_login", "render_save_status", "set_highlight_state", "add_more_questions", "add_answers", "add_gap_cards", "_render_question_only", "render_card_view", "set_interview_card_state", "render_interview_card_section", "set_question_sheet_state", "render_question_sheet_section", "current_bank", "goto_section", "render_question_bank_section", "RUN_GRADES", "_run_timer_html", "_interview_minutes", "_run_move", "_run_toggle_tail", "_run_finish", "render_run_mode", "_app_base_url", "_qr_png", "_timer_html", "_score_bar_md", "_render_feedback", "_student_api_key", "render_student_page", "_student_ctx", "_student_save", "_student_practice_tab", "render_teacher_practice", "render_practice_board", "render_class_practice_stats"]
