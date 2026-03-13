"""LangGraph state schema for PPT generation pipeline."""

from __future__ import annotations

from typing import Any, Dict, List, TypedDict


class GraphState(TypedDict, total=False):
    # Input
    source_path: str

    # Extracted full text
    extracted_text: str

    # Split result
    # sections: [{"title": "<섹션명>", "text": "<섹션 텍스트>"} ...]
    sections: List[Dict[str, str]]
    section_chunks: Dict[str, str]

    # Gemini result by section
    # section_decks[section] = {"section":..., "deck_title":..., "slides":[...]}
    section_decks: Dict[str, Any]

    # Merge result
    deck_json: Dict[str, Any]
    deck_title: str

    # Output options
    output_dir: str
    output_filename: str
    render_mode: str
    template_pptx_path: str
    template_ppt_path: str

    # Gemini options
    gemini_model: str
    gemini_temperature: float
    gemini_max_output_tokens: int
    gemini_max_retries: int
    gemini_image_model: str

    # Gamma options/results
    gamma_theme: str
    gamma_timeout_sec: int
    gamma_generation_id: str
    gamma_result: Dict[str, Any]
    pptx_url: str
    pptx_path: str

    # Final result
    final_ppt_path: str

    # Optional postprocess options
    font_name: str
    force_rewrite_agenda: bool
    save_checkpoint: bool
    enable_gemini_diagram_images: bool
    gemini_image_max_count: int
    gemini_cover_image_only: bool
    min_slide_count: int
    postprocess_rewrite_cover: bool
    postprocess_rewrite_agenda: bool
    postprocess_style_tables: bool
    postprocess_trim_ending: bool
    postprocess_apply_template: bool
    postprocess_apply_background_image: bool
    postprocess_background_image_path: str
    postprocess_background_profile: str
    postprocess_background_base_dir: str
    postprocess_background_random_seed: int
    postprocess_remove_background_image: bool


def create_empty_state() -> GraphState:
    return {
        "source_path": "",
        "extracted_text": "",
        "sections": [],
        "section_chunks": {},
        "section_decks": {},
        "deck_json": {},
        "deck_title": "",
        "final_ppt_path": "",
    }
