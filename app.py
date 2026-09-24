import os
import re
import io
import base64
from urllib.parse import urlparse, parse_qs
import streamlit as st
from dotenv import load_dotenv
from youtube_transcript_api import YouTubeTranscriptApi
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Image as ReportLabImage
)

from pypdf import PdfReader, PdfWriter

from pptx import Presentation
from pptx.util import Inches, Pt

from langchain_core.documents import Document
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_openai import OpenAIEmbeddings, ChatOpenAI
from langchain_community.vectorstores import Chroma
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.output_parsers import StrOutputParser
from langchain_core.runnables import RunnablePassthrough


# ============================================================
# 1. LOAD ENVIRONMENT VARIABLES
# ============================================================

load_dotenv()

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ============================================================
# PRESENTATION ASSETS
# ============================================================

PRESENTATION_ASSET_DIR = os.path.join(
    BASE_DIR,
    "assets",
    "presentation"
)

PRESENTATION_DASHBOARD_IMAGE = os.path.join(
    PRESENTATION_ASSET_DIR,
    "Dashboard.png"
)

PRESENTATION_CHAT_1_IMAGE = os.path.join(
    PRESENTATION_ASSET_DIR,
    "Chat-1.png"
)

PRESENTATION_CHAT_2_IMAGE = os.path.join(
    PRESENTATION_ASSET_DIR,
    "Chat-2.png"
)

PRESENTATION_WORKFLOW_IMAGE = os.path.join(
    PRESENTATION_ASSET_DIR,
    "Workflow.png"
)
# ============================================================
# FULL-PAGE YOUTUBE BACKGROUND
# ============================================================
# Place the generated background image here:
#
#     assets/banner.png
#
# The application uses this local image as the ONLY page background.
# This keeps the design stable and avoids an unrelated remote fallback.
YOUTUBE_THEME_BACKGROUND = os.path.join(
    BASE_DIR,
    "assets",
    "banner.png"
)

# Developer information comes ONLY from .env
DEVELOPER_NAME = os.getenv(
    "DEVELOPER_NAME",
    ""
)

DEVELOPER_EMAIL = os.getenv(
    "DEVELOPER_EMAIL",
    ""
)

DEVELOPER_PHONE = os.getenv(
    "DEVELOPER_PHONE",
    ""
)

# ============================================================
# 2. PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="YouTube AI Assistant",
    page_icon="🎥",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ============================================================
# 3. PREMIUM UI
# ============================================================

st.html("""
<style>

@import url(
'https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap'
);

* {
    font-family: 'Inter', sans-serif;
}

html,
body,
.stApp {
    min-height: 100%;
}

.stApp {
    background: transparent !important;
    min-height: 100vh;
    color: #f8fafc;
}

html,
body {

    background:
        transparent !important;

    min-height: 100%;
}


[data-testid="stAppViewContainer"] {

    background:
        transparent !important;
}


[data-testid="stHeader"] {

    background:
        transparent !important;
}


.main {

    background:
        transparent !important;
}


.main .block-container {

    background:
        transparent !important;
}



.hero,
.empty-state,
.premium-card,
.stat-card,
.support-card {

    background:
        rgba(255, 255, 255, 0.06) !important;

    backdrop-filter:
        blur(8px);

    -webkit-backdrop-filter:
        blur(8px);

    border:
        1px solid
        rgba(255, 255, 255, 0.16);

    box-shadow:
        0 18px 50px
        rgba(0, 0, 0, 0.25);
}


/* ============================================================
   MAIN CONTAINER
   ============================================================ */

.main .block-container {

    max-width:
        1088px !important;

    width:
        100% !important;

    padding-top:
        0 !important;

    padding-bottom:
        2rem !important;

    margin:
        0 auto !important;
}


/* ============================================================
   HIDE STREAMLIT DEFAULT ELEMENTS
   ============================================================ */

#MainMenu {
    visibility: hidden;
}

footer {
    visibility: hidden;
}

[data-testid="stToolbar"] {
    visibility: hidden;
}

[data-testid="stSidebarCollapseButton"] {
    display: none !important;
}

header {
    background: transparent !important;
}

/* ============================================================
   YOUTUBE AI SIDEBAR — SCREENSHOT MATCH
   ============================================================ */

[data-testid="stSidebar"] {
    min-width: 290px !important;
    max-width: 290px !important;
    width: 290px !important;

    background:
        linear-gradient(
            180deg,
            rgba(5, 23, 45, 0.97) 0%,
            rgba(10, 25, 48, 0.97) 38%,
            rgba(39, 15, 39, 0.98) 72%,
            rgba(62, 9, 38, 0.98) 100%
        ) !important;

    border-right:
        1px solid rgba(255, 255, 255, 0.08) !important;

    box-shadow:
        8px 0 35px rgba(0, 0, 0, 0.28) !important;
}


[data-testid="stSidebar"] > div:first-child {
    width: 290px !important;

    padding:
        0 21px 0 21px !important;

    margin-top: 0 !important;

    background:
        transparent !important;
}


/* Remove unnecessary Streamlit spacing */

[data-testid="stSidebar"] .block-container,
[data-testid="stSidebarContent"],
[data-testid="stSidebarUserContent"] {
    padding-top: 0 !important;
    margin-top: 0 !important;
    padding-bottom: 0 !important;
}

/* ============================================================
   SIDEBAR BRAND
   ============================================================ */

.sidebar-brand {
    padding: 0 5px 19px 5px;

    margin: 0;
}


[data-testid="stSidebar"] section {
    padding-top: 0 !important;
}

[data-testid="stSidebar"] [data-testid="stSidebarContent"] {
    padding-top: 0 !important;
}

[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] {
    padding-top: 0 !important;
}


.sidebar-brand-icon {
    width: 70px;
    height: 48px;

    display: flex;
    align-items: center;
    justify-content: center;

    background: #ff0000;

    border-radius: 13px;

    color: #ffffff;

    font-size: 26px;

    margin-bottom: 13px;

    box-shadow:
        0 8px 22px rgba(255, 0, 0, 0.28);
}

.sidebar-brand-icon .youtube-play-triangle {
    font-size: 24px;
    line-height: 1;
    margin-left: 2px;
}


.sidebar-brand-title {
    color: #ffffff;

    font-size: 27px;

    line-height: 1.1;

    font-weight: 800;

    letter-spacing: -1px;

    margin: 0;
}


.sidebar-brand-text {
    color: #a9c1df;

    font-size: 15px;

    line-height: 1.45;

    margin-top: 9px;

    max-width: 255px;
}


/* ============================================================
   SIDEBAR SECTION HEADINGS
   ============================================================ */

.sidebar-section {
    color: #a9c4e5;

    font-size: 13px;

    font-weight: 700;

    letter-spacing: 0.7px;

    text-transform: uppercase;

    margin:
        18px 0 9px 2px;
}


/* ============================================================
   YOUTUBE INPUT
   ============================================================ */

[data-testid="stSidebar"] .stTextInput {
    margin-top: 0 !important;
    margin-bottom: 4px !important;
}


[data-testid="stSidebar"] .stTextInput > label {
    display: none !important;
}


[data-testid="stSidebar"] .stTextInput > div > div > input {
    height: 43px !important;

    background:
        rgba(24, 42, 72, 0.92) !important;

    color:
        #ffffff !important;

    border:
        1px solid rgba(126, 160, 204, 0.35) !important;

    border-radius:
        10px !important;

    padding:
        0 13px !important;

    font-size:
        13px !important;

    box-shadow:
        inset 0 1px 0 rgba(255,255,255,0.03) !important;
}


[data-testid="stSidebar"] .stTextInput > div > div > input::placeholder {
    color: #8fa8c8 !important;
    opacity: 1 !important;
}


[data-testid="stSidebar"] .stTextInput > div > div > input:focus {
    border-color:
        rgba(125, 211, 252, 0.65) !important;

    box-shadow:
        0 0 0 2px rgba(56,189,248,0.08) !important;
}


/* ============================================================
   INPUT DESCRIPTION
   ============================================================ */

[data-testid="stSidebar"] [data-testid="stCaptionContainer"] {
    color: #9db6d6 !important;

    font-size: 12px !important;

    line-height: 1.45 !important;

    margin:
        4px 2px 13px 2px !important;
}


/* ============================================================
   PROCESS VIDEO BUTTON
   ============================================================ */

[data-testid="stSidebar"] .stButton > button {
    height: 45px !important;

    width: 100% !important;

    border:
        none !important;

    border-radius:
        10px !important;

    background:
        linear-gradient(
            135deg,
            #ff1515 0%,
            #ff202d 50%,
            #ff1010 100%
        ) !important;

    color:
        #ffffff !important;

    font-size:
        15px !important;

    font-weight:
        700 !important;

    box-shadow:
        0 9px 22px rgba(255, 0, 25, 0.30) !important;

    margin:
        0 0 15px 0 !important;
}


[data-testid="stSidebar"] .stButton > button:hover {
    transform: translateY(-1px) !important;

    box-shadow:
        0 12px 27px rgba(255, 0, 25, 0.42) !important;
}


/* ============================================================
   SIDEBAR DIVIDER
   ============================================================ */

.sidebar-divider {
    height: 1px;

    width: 100%;

    background:
        rgba(255,255,255,0.12);

    margin:
        10px 0 17px 0;
}


/* ============================================================
   HOW IT WORKS / TECHNOLOGY
   ============================================================ */

.sidebar-list {
    color: #b9cce5;

    font-size: 15px;

    line-height: 1.52;
}


.sidebar-list div {
    margin-bottom: 2px;
}


.sidebar-number {
    display: inline-flex;

    align-items: center;

    justify-content: center;

    width: 22px;

    height: 22px;

    margin-right: 8px;

    border:
        1px solid rgba(190, 215, 245, 0.55);

    border-radius: 50%;

    color: #d9e8fa;

    font-size: 12px;

    font-weight: 600;
}


.sidebar-bullet {
    margin-right: 8px;

    color: #d5e3f4;
}


/* ============================================================
   SUPPORT CARD
   ============================================================ */

.sidebar-support {
    margin-top: 7px;

    padding:
        10px 12px;

    border:
        1px solid rgba(255, 70, 105, 0.40);

    border-radius:
        8px;

    background:
        rgba(41, 14, 38, 0.58);

    box-shadow:
        inset 0 1px 0 rgba(255,255,255,0.025);
}


.sidebar-support-row {
    display: flex;

    align-items: flex-start;

    gap: 10px;
}


.sidebar-support-icon {
    font-size: 25px;

    line-height: 1;

    margin-top: 1px;
}


.sidebar-support-title {
    color: #ffffff;

    font-size: 14px;

    font-weight: 700;

    line-height: 1.2;
}


.sidebar-support-subtitle {
    color: #b6c9e0;

    font-size: 11px;

    margin-top: 2px;
}


.sidebar-contact {
    margin-top: 8px;

    color: #a9c0dc;

    font-size: 11px;

    line-height: 1.65;
}


.sidebar-contact div {
    white-space: nowrap;
}


/* ============================================================
   SIDEBAR FOOTER
   ============================================================ */

.sidebar-footer {
    color: #9bb0cb;

    font-size: 11px;

    line-height: 1.45;

    margin:
        9px 0 0 2px;
}


.sidebar-footer strong {
    color: #dce9f8;

    font-size: 12px;
}


.sidebar-heart {
    color: #ff2342;

    font-size: 17px;

    margin-left: 5px;
}


/* ============================================================
   REMOVE SIDEBAR SCROLLBAR VISUAL
   ============================================================ */

[data-testid="stSidebar"]::-webkit-scrollbar {
    width: 3px;
}

[data-testid="stSidebar"]::-webkit-scrollbar-track {
    background: transparent;
}

[data-testid="stSidebar"]::-webkit-scrollbar-thumb {
    background: rgba(255,255,255,0.10);

    border-radius: 10px;
}

/* ============================================================
   HERO
   ============================================================ */

.hero {

    position:
        relative;

    overflow:
        hidden;

    /* Lighter hero overlay */
    background:
        rgba(20, 32, 55, 0.28) !important;

    backdrop-filter:
        blur(5px);

    -webkit-backdrop-filter:
        blur(5px);

    border:
        1px solid
        rgba(255, 255, 255, 0.22);

    border-radius:
        20px;

    box-shadow:
        0 20px 60px
        rgba(0, 0, 0, 0.20);

    min-height:
        336px;

    padding:
        34px 54px 28px 54px;

    margin-bottom:
        20px !important;
}


.hero::before {

    content: "";
    position: absolute;
    width: 360px;
    height: 360px;
    top: -190px;
    right: -90px;
    border-radius: 50%;
    background:
        rgba(96, 165, 250, 0.12);

    filter: blur(80px);
    pointer-events: none;
}


.hero::after {

    content: "";
    position: absolute;
    width: 280px;
    height: 280px;
    bottom: -180px;
    left: -90px;
    border-radius: 50%;
    background:
        rgba(244, 114, 182, 0.08);

    filter: blur(70px);
    pointer-events: none;
}


/* ============================================================
   HERO BADGE
   ============================================================ */

.hero-badge {

    position: relative;

    display: inline-flex;

    align-items: center;

    padding:
        7px
        14px;

    border-radius: 999px;

    background:
        rgba(255, 255, 255, 0.08);

    border:
        1px solid rgba(191, 219, 254, 0.25);

    color: #dbeafe;

    font-size: 12px;

    font-weight: 750;

    margin-bottom: 18px;
}


/* ============================================================
   HERO TITLE
   ============================================================ */
.hero-title {

    position:
        relative;

    display:
        flex;

    align-items:
        baseline;

    flex-wrap:
        wrap;

    gap:
        12px;

    font-size:
        clamp(42px, 4vw, 58px);

    line-height:
        1.05;

    font-weight:
        850;

    letter-spacing:
        -2.2px;

    color:
        #ffffff;

    margin:
        0;

    text-shadow:
        0 4px 18px
        rgba(0, 0, 0, 0.38);
}

.hero-gradient {

    color:
        #FF1A30 !important;

    background:
        #FF1A30 !important;

    -webkit-background-clip:
        text !important;

    background-clip:
        text !important;

    -webkit-text-fill-color:
        #FF1A30 !important;

    text-shadow:
        0 3px 10px
        rgba(255, 26, 48, 0.25);

    font-weight:
        900;

    letter-spacing:
        -2.4px;
}


.hero-main-title {

    color:
        #ffffff;

    font-weight:
        850;

    letter-spacing:
        -2.2px;

    text-shadow:
        0 4px 16px
        rgba(0, 0, 0, 0.40);
}

/* ============================================================
   HERO TITLE — RESPONSIVE
   ============================================================ */

@media (max-width: 700px) {

    .hero-title {

        font-size:
            36px;

        gap:
            7px;

        letter-spacing:
            -1.5px;
    }

    .hero-gradient,
    .hero-main-title {

        letter-spacing:
            -1.5px;
    }
}

.hero-description {

    position: relative;
    max-width: 980px;
    color: #f8fafc;
    font-size: 16px;
    line-height: 1.55;
    font-weight: 500;
    margin-top: 16px;
    text-shadow:
        0 2px 10px rgba(0, 0, 0, 0.45);
}

.hero-features {
    position: relative;

    display: grid;

    grid-template-columns:
        repeat(3, 1fr);

    gap: 35px;

    margin-top: 22px;
}


.hero-feature {
    display: flex;

    align-items: center;

    gap: 15px;
}


.hero-feature-icon {
    width: 64px;
    height: 64px;

    flex-shrink: 0;

    display: flex;

    align-items: center;
    justify-content: center;

    border-radius: 50%;

    background:
        rgba(55, 83, 140, 0.45);

    border:
        1px solid
        rgba(125, 170, 245, 0.25);

    color: #ffffff;

    font-size: 28px;

    box-shadow:
        inset 0 1px 0 rgba(255,255,255,0.08),
        0 8px 25px rgba(0,0,0,0.18);
}


.hero-feature:nth-child(3)
.hero-feature-icon {
    background:
        rgba(126, 20, 92, 0.55);
}


.hero-feature-title {
    color: #ffffff;

    font-size: 15px;

    font-weight: 800;

    line-height: 1.2;

    margin-bottom: 5px;
}


.hero-feature-text {

    color: #ffffff;
    font-size: 13px;
    line-height: 1.4;
    font-weight: 500;
    opacity: 0.92;
    text-shadow:
        0 2px 8px rgba(0, 0, 0, 0.5);
}


@media (max-width: 850px) {

    .hero-features {
        grid-template-columns: 1fr;

        gap: 18px;
    }

}

/* ============================================================
   PREMIUM CARDS
   ============================================================ */

.premium-card {

    background: rgba(255, 255, 255, 0.07);

    backdrop-filter:
        blur(10px);

    -webkit-backdrop-filter:
        blur(10px);

    border:
        1px solid rgba(255, 255, 255, 0.11);

    border-radius: 22px;

    padding: 24px;

    box-shadow:
        0 20px 55px
        rgba(0, 0, 0, 0.23);
}


.section-title {

    font-size: 21px;

    font-weight: 800;

    color: #ffffff;

    margin-bottom: 5px;
}


.section-subtitle {

    font-size: 13px;

    color: #a9bad1;

    margin-bottom: 18px;
}


/* ============================================================
   STAT CARDS
   ============================================================ */

.stat-card {

    background: rgba(255, 255, 255, 0.07);

    backdrop-filter:
        blur(10px);

    -webkit-backdrop-filter:
        blur(10px);

    border:
        1px solid rgba(255, 255, 255, 0.11);

    border-radius: 19px;

    padding: 20px;

    min-height: 108px;

    box-shadow:
        0 15px 40px
        rgba(0, 0, 0, 0.20);
}


.stat-icon {

    font-size: 21px;

    margin-bottom: 8px;
}


.stat-label {

    color: #9db2ce;

    font-size: 11px;

    font-weight: 800;

    text-transform: uppercase;

    letter-spacing: 0.7px;
}


.stat-value {

    color: #ffffff;

    font-size: 18px;

    font-weight: 750;

    margin-top: 4px;

    word-break: break-word;
}


/* ============================================================
   INPUT
   ============================================================ */

.stTextInput > div > div > input {

    background:
        rgba(8, 25, 50, 0.82) !important;

    color:
        #ffffff !important;

    border:
        1px solid rgba(191, 219, 254, 0.18) !important;

    border-radius:
        13px !important;

    padding:
        13px 15px !important;

    font-size:
        14px !important;
}


.stTextInput > div > div > input:focus {

    border-color:
        rgba(125, 211, 252, 0.75) !important;

    box-shadow:
        0 0 0 2px
        rgba(56, 189, 248, 0.12) !important;
}


/* ============================================================
   BUTTON
   ============================================================ */

.stButton > button {

    width: 100%;

    border:
        1px solid rgba(255, 255, 255, 0.16);

    border-radius:
        13px;

    padding:
        12px 18px;

    font-weight:
        750;

    color:
        #ffffff;

    background:
        linear-gradient(
            135deg,
            #ff0000,
            #ff1f1f,
            #cc0000
        );

    box-shadow:
        0 12px 30px
        rgba(255, 0, 0, 0.30);

    transition:
        transform 0.18s ease,
        box-shadow 0.18s ease;
}


.stButton > button:hover {

    transform:
        translateY(-2px);

    box-shadow:
        0 17px 38px
        rgba(255, 0, 0, 0.48);
}


/* ============================================================
   CHAT
   ============================================================ */

[data-testid="stChatMessage"] {

    background:
        rgba(15, 35, 65, 0.38);

    border:
        1px solid rgba(191, 219, 254, 0.10);

    border-radius:
        18px;

    margin-bottom:
        10px;
}


[data-testid="stChatInput"] {

    border-radius:
        16px;
}

/* ============================================================
   CONVERSATION PDF ICON
   ============================================================ */
[data-testid="stDownloadButton"] {
    width: 100% !important;
    margin: 14px 0 8px 0 !important;
}

[data-testid="stDownloadButton"] button {
    width: 100% !important;
    min-height: 44px !important;

    padding: 10px 18px !important;

    display: flex !important;
    align-items: center !important;
    justify-content: center !important;

    border-radius: 12px !important;

    border:
        1px solid
        rgba(255, 255, 255, 0.18) !important;

    background:
        rgba(20, 30, 52, 0.88) !important;

    color: #ffffff !important;

    font-size: 14px !important;
    font-weight: 700 !important;

    box-shadow:
        0 8px 22px
        rgba(0, 0, 0, 0.22) !important;
}

[data-testid="stDownloadButton"] button:hover {
    background:
        rgba(40, 52, 78, 0.96) !important;

    border-color:
        rgba(255, 255, 255, 0.35) !important;

    transform:
        translateY(-1px) !important;
}

    /* ============================================================
    SIDEBAR PROJECT PRESENTATION
    SAME DESIGN AS PROCESS VIDEO
    ============================================================ */

    [data-testid="stSidebar"]
    [data-testid="stDownloadButton"] {

        width: 100% !important;

        margin:
            0 0 15px 0 !important;

        padding:
            0 !important;
    }


   
    [data-testid="stSidebar"] [data-testid="stDownloadButton"] button {

        width: 100% !important;

        min-height:
            45px !important;

        height:
            auto !important;

        padding:
            10px 12px !important;

        border:
            none !important;

        border-radius:
            10px !important;

        background:
            linear-gradient(
                135deg,
                #ff1515 0%,
                #ff202d 50%,
                #ff1010 100%
            ) !important;

        color:
            #ffffff !important;

        font-size:
            13px !important;

        font-weight:
            700 !important;

        line-height:
            1.35 !important;

        white-space:
            normal !important;

        text-align:
            center !important;

        box-shadow:
            0 9px 22px
            rgba(255, 0, 25, 0.30) !important;

        transition:
            transform 0.18s ease,
            box-shadow 0.18s ease !important;
    }


    [data-testid="stSidebar"]
    [data-testid="stDownloadButton"] button:hover {

        transform:
            translateY(-1px) !important;

        box-shadow:
            0 12px 27px
            rgba(255, 0, 25, 0.42) !important;
    }


/* ============================================================
   STATUS
   ============================================================ */

.status-pill {

    display: inline-flex;

    align-items: center;

    gap: 7px;

    padding:
        7px 12px;

    border-radius:
        999px;

    background:
        rgba(34, 197, 94, 0.10);

    border:
        1px solid
        rgba(34, 197, 94, 0.25);

    color:
        #bbf7d0;

    font-size:
        11px;

    font-weight:
        800;
}


.status-dot {

    width:
        7px;

    height:
        7px;

    border-radius:
        50%;

    background:
        #4ade80;

    box-shadow:
        0 0 11px
        rgba(74, 222, 128, 0.9);
}


/* ============================================================
   PREMIUM EMPTY STATE
   ============================================================ */

.empty-state {

    position: relative;

    min-height: 300px;

    padding: 42px 50px 36px 50px;

    display: flex;

    flex-direction: column;

    align-items: center;

    justify-content: center;

    text-align: center;

    background:
        linear-gradient(
            145deg,
            rgba(10, 19, 40, 0.72),
            rgba(21, 20, 43, 0.58)
        );

    backdrop-filter:
        blur(12px);

    -webkit-backdrop-filter:
        blur(12px);

    border:
        1px solid
        rgba(160, 196, 239, 0.30);

    border-radius:
        20px;

    box-shadow:
        0 22px 55px
        rgba(0, 0, 0, 0.24);

    overflow:
        hidden;
}


/* subtle premium glow */

.empty-state::before {

    content: "";

    position: absolute;

    width: 280px;

    height: 180px;

    top: -100px;

    left: 50%;

    transform:
        translateX(-50%);

    background:
        rgba(255, 30, 55, 0.13);

    filter:
        blur(65px);

    pointer-events:
        none;
}


/* ============================================================
   ICON
   ============================================================ */

.empty-icon {

    position: relative;

    width: 52px;

    height: 52px;

    display: flex;

    align-items: center;

    justify-content: center;

    margin-bottom:
        14px;

    border-radius:
        16px;

    background:
        rgba(255, 255, 255, 0.08);

    border:
        1px solid
        rgba(255, 255, 255, 0.14);

    font-size:
        27px;

    box-shadow:
        0 10px 25px
        rgba(0, 0, 0, 0.18);
}


/* ============================================================
   TITLE
   ============================================================ */

.empty-title {

    position: relative;

    margin:
        0 0 9px 0;

    color:
        #ffffff;

    font-size:
        26px;

    line-height:
        1.2;

    font-weight:
        800;

    letter-spacing:
        -0.6px;

    text-shadow:
        0 3px 12px
        rgba(0, 0, 0, 0.40);
}


/* ============================================================
   DESCRIPTION
   ============================================================ */

.empty-text {

    position: relative;

    max-width:
        820px;

    margin:
        0 auto;

    color:
        rgba(241, 245, 249, 0.88);

    font-size:
        15px;

    line-height:
        1.65;

    font-weight:
        450;

    text-shadow:
        0 2px 8px
        rgba(0, 0, 0, 0.35);
}

.empty-text strong {

    color:
        #ffffff;

    font-weight:
        700;
}


/* ============================================================
   PREMIUM QUOTE
   ============================================================ */

.empty-quote {

    position: relative;

    width:
        min(620px, 92%);

    margin:
        22px auto 0 auto;

    padding:
        13px 22px;

    display:
        flex;

    align-items:
        center;

    justify-content:
        center;

    gap:
        10px;

    box-sizing:
        border-box;

    border:
        1px solid
        rgba(100, 160, 235, 0.38);

    border-radius:
        999px;

    background:
        rgba(4, 18, 40, 0.68);

    color:
        rgba(248, 250, 252, 0.90);

    font-size:
        13px;

    line-height:
        1.45;

    font-weight:
        600;

    font-style:
        italic;

    box-shadow:
        inset 0 1px 0
        rgba(255,255,255,0.05),

        0 10px 28px
        rgba(0,0,0,0.16);
}


.empty-quote-icon {

    flex-shrink:
        0;

    margin:
        0;

    font-size:
        18px;

    font-style:
        normal;
}


/* ============================================================
   MOBILE
   ============================================================ */

@media (max-width: 700px) {

    .empty-state {

        min-height:
            280px;

        padding:
            32px 22px;
    }

    .empty-title {

        font-size:
            23px;
    }

    .empty-text {

        font-size:
            14px;
    }

    .empty-quote {

        width:
            100%;

        border-radius:
            16px;

        padding:
            12px 15px;
    }
}


/* ============================================================
   API / SUPPORT CARD
   ============================================================ */

.support-card {

    background:

        linear-gradient(
            135deg,
            rgba(127, 29, 29, 0.68),
            rgba(88, 28, 135, 0.65)
        );

    border:
        1px solid
        rgba(251, 191, 36, 0.25);

    border-radius:
        22px;

    padding:
        25px;

    margin:
        15px 0 25px 0;

    box-shadow:
        0 20px 55px
        rgba(0, 0, 0, 0.24);
}


.support-title {

    color:
        #ffffff;

    font-size:
        20px;

    font-weight:
        800;

    margin-bottom:
        8px;
}


.support-text {

    color:
        #e2e8f0;

    font-size:
        14px;

    line-height:
        1.7;
}


.support-contact {

    margin-top:
        15px;

    padding:
        15px;

    border-radius:
        14px;

    background:
        rgba(0, 0, 0, 0.20);

    color:
        #ffffff;

    line-height:
        1.9;
}


/* ============================================================
   PROCESSING CARD
   ============================================================ */

.processing-card {

    background:

        linear-gradient(
            135deg,
            rgba(22, 78, 99, 0.78),
            rgba(55, 48, 107, 0.75)
        );

    border:
        1px solid
        rgba(125, 211, 252, 0.18);

    border-radius:
        20px;

    padding:
        20px;

    margin-bottom:
        20px;
}


/* ============================================================
   FOOTER
   ============================================================ */

.custom-footer {

    text-align:
        center;

    margin-top:
        48px;

    padding-top:
        22px;

    border-top:
        1px solid
        rgba(255, 255, 255, 0.10);

    color:
        #8ea5c7;

    font-size:
        12px;

    line-height:
        1.8;
}


.custom-footer strong {

    color:
        #dbeafe;
}

/* ============================================================
   YOUTUBE FOOTER
   ============================================================ */

.youtube-footer {

    text-align:
        center;

    margin-top:
        42px;

    padding-bottom:
        5px;
}


.youtube-footer-links {

    color:
        #d8dfec;

    font-size:
        11px;

    font-weight:
        600;

    letter-spacing:
        5px;

    line-height:
        1.6;

    text-transform:
        uppercase;

    white-space:
        nowrap;
}


.youtube-footer-links span {

    color:
        #d5dce8;

    margin:
        0 8px;

    letter-spacing:
        1px;
}


.youtube-footer-line {

    width:
        48px;

    height:
        3px;

    margin:
        12px auto 0 auto;

    border-radius:
        999px;

    background:
        #ff1744;

    box-shadow:
        0 0 10px
        rgba(255, 23, 68, 0.55);
}


@media (max-width: 700px) {

    .youtube-footer-links {

        letter-spacing:
            2px;

        font-size:
            9px;
    }

    .youtube-footer-links span {

        margin:
            0 3px;
    }
}


/* ============================================================
   CAPTION
   ============================================================ */

[data-testid="stCaptionContainer"] {

    color:
        #9db2ce !important;
}


/* ============================================================
   STATUS WIDGET
   ============================================================ */

[data-testid="stStatusWidget"] {

    background:
        linear-gradient(
            135deg,
            rgba(22, 54, 91, 0.96),
            rgba(56, 38, 88, 0.96)
        );

    border:
        1px solid
        rgba(255, 255, 255, 0.12);

    border-radius:
        18px;
}


/* ============================================================
   EXPANDER
   ============================================================ */

[data-testid="stExpander"] {

    background:
        rgba(15, 35, 65, 0.65);

    border:
        1px solid
        rgba(255, 255, 255, 0.10);

    border-radius:
        15px;
}

/* ============================================================
   REFERENCE SCREENSHOT SPACING
   ============================================================ */

.hero {
    margin-bottom:
        16px !important;
}

.empty-state {
    margin-top:
        0 !important;

    margin-bottom:
        0 !important;
}

.premium-card {
    margin-top:
        0 !important;
}

.stat-card {
    margin-top:
        0 !important;
}

.stVerticalBlock {
    gap:
        0 !important;
}

/* ============================================================
   PROJECT PRESENTATION
   ============================================================ */

.presentation-wrapper {
    margin-top: 28px;
}

.presentation-hero {
    background:
        linear-gradient(
            135deg,
            rgba(7, 20, 43, 0.92),
            rgba(44, 14, 49, 0.88)
        );
    border: 1px solid rgba(255,255,255,0.16);
    border-radius: 20px;
    padding: 28px;
    margin-bottom: 18px;
    box-shadow: 0 18px 45px rgba(0,0,0,0.24);
}

.presentation-title {
    color: #ffffff;
    font-size: 28px;
    font-weight: 800;
    margin-bottom: 8px;
}

.presentation-subtitle {
    color: #c7d7ed;
    font-size: 14px;
    line-height: 1.65;
}

.presentation-flow {
    margin-top: 20px;
    padding: 15px 18px;
    border-radius: 14px;
    background: rgba(255,255,255,0.06);
    border: 1px solid rgba(255,255,255,0.10);
    color: #e2e8f0;
    font-size: 13px;
    line-height: 1.7;
    text-align: center;
}

.presentation-grid {
    display: grid;
    grid-template-columns: repeat(2, minmax(0, 1fr));
    gap: 16px;
    margin-top: 18px;
}

.presentation-card {
    background: rgba(7,19,40,0.82);
    border: 1px solid rgba(255,255,255,0.12);
    border-radius: 17px;
    padding: 20px;
    box-shadow: 0 12px 30px rgba(0,0,0,0.20);
}

.presentation-card h3 {
    color: #ffffff;
    font-size: 16px;
    margin: 0 0 9px 0;
}

.presentation-card p,
.presentation-card li {
    color: #cbd8ea;
    font-size: 13px;
    line-height: 1.65;
}

.presentation-card ul {
    margin: 7px 0 0 18px;
    padding: 0;
}

.presentation-screenshot-title {
    color: #ffffff;
    font-size: 21px;
    font-weight: 800;
    margin: 30px 0 14px 0;
}

.presentation-screenshot {
    background: rgba(7,19,40,0.82);
    border: 1px solid rgba(255,255,255,0.14);
    border-radius: 18px;
    padding: 10px;
    margin-bottom: 18px;
    box-shadow: 0 16px 40px rgba(0,0,0,0.25);
}

.presentation-screenshot img {
    width: 100%;
    border-radius: 12px;
    display: block;
}

.presentation-screenshot-caption {
    color: #aebfd7;
    font-size: 12px;
    text-align: center;
    padding: 10px 5px 4px 5px;
}

.presentation-download {
    margin-top: 22px;
}

@media (max-width: 800px) {
    .presentation-grid {
        grid-template-columns: 1fr;
    }
}

</style>
""")


# ============================================================
# 4. SESSION STATE
# ============================================================

if "rag_chain" not in st.session_state:
    st.session_state.rag_chain = None

if "video_processed" not in st.session_state:
    st.session_state.video_processed = False

if "video_id" not in st.session_state:
    st.session_state.video_id = ""

if "chunk_count" not in st.session_state:
    st.session_state.chunk_count = 0

if "transcript_text" not in st.session_state:
    st.session_state.transcript_text = ""

if "chat_history" not in st.session_state:
    st.session_state.chat_history = []

if "api_error" not in st.session_state:
    st.session_state.api_error = False

if "api_error_type" not in st.session_state:
    st.session_state.api_error_type = ""


# ============================================================
# 5. HELPER FUNCTIONS
# ============================================================

def developer_contact_html():
    """
    Creates the developer contact information.
    Values are loaded from .env.
    """

    email_html = ""

    phone_html = ""

    if DEVELOPER_EMAIL:
        email_html = f"""
        <div>
            ✉️ <strong>Email:</strong>
            {DEVELOPER_EMAIL}
        </div>
        """

    if DEVELOPER_PHONE:
        phone_html = f"""
        <div>
            ☎️ <strong>Phone:</strong>
            {DEVELOPER_PHONE}
        </div>
        """

    return f"""
    <div class="support-contact">

        <div>
            👨‍💻 <strong>Developer:</strong>
            {DEVELOPER_NAME}
        </div>

        {email_html}

        {phone_html}

    </div>
    """


def show_api_support_message(error_type="invalid"):
    """
    Displays a friendly API-related message.
    Never exposes the raw OpenAI API key/error.
    """

    if error_type == "missing":

        title = "🔐 OpenAI connection is not configured"

        message = (
            "The application is ready, but the OpenAI API "
            "connection has not been configured yet."
        )

    elif error_type == "invalid":

        title = "🔑 OpenAI API key needs attention"

        message = (
            "Your video transcript was found successfully, "
            "but the AI service could not connect using the "
            "current OpenAI API key."
        )

    else:

        title = "⚠️ AI service temporarily unavailable"

        message = (
            "The video was found, but the AI service could "
            "not complete the request."
        )

    st.html(f"""
    <div class="support-card">

        <div class="support-title">
            {title}
        </div>

        <div class="support-text">
            {message}
            <br><br>
            Please contact the developer for assistance
            with the API configuration.
        </div>

        {developer_contact_html()}

    </div>
    """)


def is_api_error(error):
    """
    Detect common OpenAI authentication/configuration errors.
    """

    error_text = str(error).lower()

    api_error_patterns = [
        "incorrect api key",
        "invalid api key",
        "authenticationerror",
        "error code: 401",
        "401",
        "api key",
        "unauthorized"
    ]

    return any(
        pattern in error_text
        for pattern in api_error_patterns
    )

def format_docs(documents):

    return "\n\n".join(
        document.page_content
        for document in documents
    )

def is_greeting(question):
    """Return True when the user is only greeting the assistant."""
    normalized = re.sub(r"[^a-z\s]", "", question.lower()).strip()

    greeting_patterns = {
        "hi",
        "hello",
        "hey",
        "good morning",
        "good afternoon",
        "good evening",
        "how are you",
        "who are you",
        "what are you"
    }

    return normalized in greeting_patterns


def get_greeting_response():
    return (
        "Hi! I’m your YouTube AI Assistant. I can answer questions "
        "about the processed video using its transcript. Ask me about "
        "the topic, people mentioned, key points, events, or a summary, "
        "and I’ll answer you."
    )

def create_conversation_pdf(chat_history, video_id):

    buffer = io.BytesIO()

    document = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=18 * mm,
        leftMargin=18 * mm,
        topMargin=15 * mm,
        bottomMargin=15 * mm
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "ConversationTitle",
        parent=styles["Title"],
        fontSize=18,
        leading=22,
        alignment=TA_CENTER,
        spaceAfter=12
    )

    role_style = ParagraphStyle(
        "ConversationRole",
        parent=styles["Heading3"],
        fontSize=11,
        leading=14,
        spaceBefore=8,
        spaceAfter=5
    )

    message_style = ParagraphStyle(
        "ConversationMessage",
        parent=styles["BodyText"],
        fontSize=10,
        leading=15,
        spaceAfter=8
    )

    story = []

    story.append(
        Paragraph(
            "YouTube Video Conversation",
            title_style
        )
    )

    story.append(
        Paragraph(
            f"Video ID: {video_id}",
            message_style
        )
    )

    if not chat_history:

        story.append(
            Paragraph(
                "No conversation has been completed yet.",
                message_style
            )
        )

    else:

        for message in chat_history:

            role = (
                "You"
                if message["role"] == "user"
                else "AI Assistant"
            )

            content = str(
                message["content"]
            ).strip()

            content = (
                content
                .replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\n", "<br/>")
            )

            story.append(
                Paragraph(
                    role,
                    role_style
                )
            )

            story.append(
                Paragraph(
                    content,
                    message_style
                )
            )

    document.build(story)

    buffer.seek(0)

    return buffer.getvalue()

# ============================================================
# PROJECT PRESENTATION PDF
# ============================================================

def create_project_presentation_pdf():
    """
    Generate a professional project presentation PDF
    for the YouTube AI Assistant.
    """

    buffer = io.BytesIO()

    document = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=18 * mm,
        leftMargin=18 * mm,
        topMargin=15 * mm,
        bottomMargin=18 * mm
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "ProjectTitle",
        parent=styles["Title"],
        fontSize=24,
        leading=29,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#172554"),
        spaceAfter=10
    )

    subtitle_style = ParagraphStyle(
        "ProjectSubtitle",
        parent=styles["Normal"],
        fontSize=11,
        leading=16,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#475569"),
        spaceAfter=18
    )

    heading_style = ParagraphStyle(
        "ProjectHeading",
        parent=styles["Heading1"],
        fontSize=17,
        leading=21,
        textColor=colors.HexColor("#0f172a"),
        spaceBefore=10,
        spaceAfter=8
    )

    body_style = ParagraphStyle(
        "ProjectBody",
        parent=styles["BodyText"],
        fontSize=9.7,
        leading=14,
        textColor=colors.HexColor("#334155"),
        spaceAfter=7
    )

    bullet_style = ParagraphStyle(
        "ProjectBullet",
        parent=body_style,
        leftIndent=14,
        firstLineIndent=-8,
        spaceAfter=4
    )

    story = []

    # --------------------------------------------------------
    # COVER
    # --------------------------------------------------------

    story.append(
        Spacer(1, 18 * mm)
    )

    story.append(
        Paragraph(
            "YouTube AI Assistant",
            title_style
        )
    )

    story.append(
        Paragraph(
            "AI-Powered Video Intelligence using "
            "Retrieval-Augmented Generation (RAG)",
            subtitle_style
        )
    )

    story.append(
        Paragraph(
            "<b>Turn long YouTube videos into an interactive "
            "knowledge base.</b>",
            subtitle_style
        )
    )

    # --------------------------------------------------------
    # DEVELOPER CONTACT — COVER PAGE
    # --------------------------------------------------------

    developer_contact_style = ParagraphStyle(
        "DeveloperContact",
        parent=styles["Normal"],
        fontSize=9,
        leading=13,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#475569"),
        spaceAfter=12
    )

    developer_contact = []

    if DEVELOPER_NAME:
        developer_contact.append(
            f"<b>Developed by {DEVELOPER_NAME}</b>"
        )

    if DEVELOPER_EMAIL:
        developer_contact.append(
            f"Email: {DEVELOPER_EMAIL}"
        )

    if DEVELOPER_PHONE:
        developer_contact.append(
            f"Phone: {DEVELOPER_PHONE}"
        )

    story.append(
        Paragraph(
            "<br/>".join(developer_contact),
            developer_contact_style
        )
    )

 

    story.append(
        Spacer(1, 8 * mm)
    )

    
    # --------------------------------------------------------
    # PROJECT DETAILS — USER / CLIENT FRIENDLY
    # --------------------------------------------------------

    story.append(
        Paragraph(
            "1. Project Overview",
            heading_style
        )
    )

    story.append(
        Paragraph(
            "YouTube AI Assistant is an AI-powered application that "
            "turns a YouTube video's transcript into an interactive "
            "knowledge source. Instead of manually searching through "
            "a long video, users can provide a YouTube link and ask "
            "questions in natural language. The system finds the "
            "most relevant information from the transcript and "
            "generates a clear answer using Retrieval-Augmented "
            "Generation (RAG).",
            body_style
        )
    )

    # --------------------------------------------------------
        # PROJECT OVERVIEW IMAGE
        # --------------------------------------------------------
    
    if os.path.exists(PRESENTATION_DASHBOARD_IMAGE):
    
        story.append(
            Spacer(
                1,
                3 * mm
            )
        )
    
        story.append(
            Paragraph(
                "<b>Application Dashboard</b>",
                body_style
            )
        )
    
        story.append(
            ReportLabImage(
                PRESENTATION_DASHBOARD_IMAGE,
                width=165 * mm,
                height=68 * mm
            )
        )
    
        story.append(
            Spacer(
                1,
                5 * mm
            )
        )
    
      

    # --------------------------------------------------------
    # AIM
    # --------------------------------------------------------

    story.append(
        Paragraph(
            "2. Aim of the Project",
            heading_style
        )
    )

    for item in [
        "Make information inside long YouTube videos easier to find.",
        "Allow users to ask questions using normal, natural language.",
        "Reduce the need to watch an entire video to find one specific answer.",
        "Retrieve relevant transcript information before generating an AI response.",
        "Provide a simple conversational interface for exploring video-based knowledge."
    ]:
        story.append(
            Paragraph(
                "• " + item,
                bullet_style
            )
        )


    # --------------------------------------------------------
    # IMPORTANCE
    # --------------------------------------------------------

    story.append(
        Paragraph(
            "3. Why This Project Is Important",
            heading_style
        )
    )

    story.append(
        Paragraph(
            "Long-form video content is increasingly used for "
            "education, training, tutorials, interviews and "
            "knowledge sharing. Finding one specific piece of "
            "information inside a long video can be time-consuming. "
            "This project provides a faster way to explore that "
            "information through conversational question answering.",
            body_style
        )
    )

    for item in [
        "Saves time when searching for specific information.",
        "Makes video-based learning more interactive.",
        "Converts transcript content into a searchable knowledge source.",
        "Uses semantic retrieval instead of relying only on keyword matching.",
        "Keeps generated answers grounded in the retrieved video transcript."
    ]:
        story.append(
            Paragraph(
                "• " + item,
                bullet_style
            )
        )


    # --------------------------------------------------------
    # WHAT THE USER CAN DO
    # --------------------------------------------------------

    story.append(
        Paragraph(
            "4. What Users Can Do",
            heading_style
        )
    )

    for item in [
        "Paste a YouTube video link.",
        "Process the video's available transcript.",
        "Ask questions naturally through the chat interface.",
        "Find specific information from the video.",
        "Ask follow-up questions about the same video.",
        "Continue the conversation using the existing chat history.",
        "Save the conversation as a PDF."
    ]:
        story.append(
            Paragraph(
                "• " + item,
                bullet_style
            )
        )


    # --------------------------------------------------------
    # IMPLEMENTATION
    # --------------------------------------------------------

    story.append(
        Paragraph(
            "5. How the Project Is Implemented",
            heading_style
        )
    )

    for item in [
        "<b>1. YouTube Input:</b> The user provides a YouTube video URL.",
        "<b>2. Video Identification:</b> The application extracts the video ID.",
        "<b>3. Transcript Retrieval:</b> YouTubeTranscriptApi retrieves an accessible transcript.",
        "<b>4. Document Preparation:</b> The transcript is converted into a document for processing.",
        "<b>5. Text Chunking:</b> LangChain divides the transcript into smaller searchable chunks.",
        "<b>6. Embeddings:</b> OpenAI text-embedding-3-small converts transcript chunks into vector representations.",
        "<b>7. Vector Storage:</b> Chroma stores the generated embeddings for semantic search.",
        "<b>8. Retrieval:</b> Relevant transcript chunks are retrieved for each user question.",
        "<b>9. AI Generation:</b> GPT-4o-mini uses the retrieved context to generate the response.",
        "<b>10. User Experience:</b> The answer is displayed through the Streamlit conversational interface."
    ]:
        story.append(
            Paragraph(
                "• " + item,
                bullet_style
            )
        )


    # --------------------------------------------------------
    # HOW IT WORKS
    # --------------------------------------------------------

    story.append(
        Paragraph(
            "6. How It Works — In Simple Terms",
            heading_style
        )
    )

    story.append(
        Paragraph(
            "<b>YouTube Video → Transcript → Text Chunks → "
            "Embeddings → Chroma Vector Search → Relevant Context → "
            "GPT-4o-mini → User Answer</b>",
            body_style
        )
    )

    for item in [
        "<b>Input:</b> A YouTube video link.",
        "<b>Understand:</b> The application retrieves and prepares the transcript.",
        "<b>Search:</b> The system finds transcript content related to the user's question.",
        "<b>Generate:</b> GPT-4o-mini creates an answer from the retrieved context.",
        "<b>Respond:</b> The answer is shown in a conversational interface."
    ]:
        story.append(
            Paragraph(
                "• " + item,
                bullet_style
            )
        )

    # --------------------------------------------------------
    # CHAT EXPERIENCE — IMAGE 1
    # --------------------------------------------------------

    if os.path.exists(PRESENTATION_CHAT_1_IMAGE):

        story.append(
            Spacer(
                1,
                3 * mm
            )
        )

        story.append(
            Paragraph(
                "<b>Natural-Language Video Q&A</b>",
                body_style
            )
        )

        story.append(
            ReportLabImage(
                PRESENTATION_CHAT_1_IMAGE,
                width=165 * mm,
                height=72 * mm
            )
        )

        story.append(
            Spacer(
                1,
                5 * mm
            )
        )


    # --------------------------------------------------------
    # USE CASES
    # --------------------------------------------------------

    story.append(
        Paragraph(
            "7. Where This Project Can Be Used",
            heading_style
        )
    )

    for item in [
        "<b>Education:</b> Ask questions about lectures, tutorials and learning videos.",
        "<b>Corporate Training:</b> Quickly locate information inside training content.",
        "<b>Technical Learning:</b> Search technical demonstrations and tutorials.",
        "<b>Research:</b> Explore interviews, discussions and information-rich videos.",
        "<b>Documentation:</b> Find information from video-based documentation.",
        "<b>Personal Learning:</b> Interactively explore educational and informational videos."
    ]:
        story.append(
            Paragraph(
                "• " + item,
                bullet_style
            )
        )


    # --------------------------------------------------------
    # KEY BENEFITS
    # --------------------------------------------------------

    story.append(
        Paragraph(
            "8. Key Benefits",
            heading_style
        )
    )

    for item in [
        "Faster access to information.",
        "Natural-language question answering.",
        "Transcript-grounded responses.",
        "Semantic information retrieval.",
        "Interactive conversation with video content.",
        "Simple Streamlit-based user experience.",
        "Conversation history for follow-up questions.",
        "PDF export for preserving the conversation."
    ]:
        story.append(
            Paragraph(
                "• " + item,
                bullet_style
            )
        )


    # --------------------------------------------------------
    # KEY FEATURES
    # --------------------------------------------------------

    story.append(
        Paragraph(
            "9. Key Features",
            heading_style
        )
    )

    for item in [
        "YouTube URL processing.",
        "Automatic video ID extraction.",
        "Transcript retrieval.",
        "Semantic document chunking.",
        "OpenAI embeddings.",
        "Chroma vector database.",
        "Semantic retrieval using relevant transcript chunks.",
        "GPT-4o-mini answer generation.",
        "Conversation history.",
        "Friendly API error handling.",
        "Conversation PDF export.",
        "Premium Streamlit interface."
    ]:
        story.append(
            Paragraph(
                "• " + item,
                bullet_style
            )
        )


    # --------------------------------------------------------
    # CHAT EXPERIENCE — IMAGE 2
    # --------------------------------------------------------

    if os.path.exists(PRESENTATION_CHAT_2_IMAGE):

        story.append(
            Spacer(
                1,
                3 * mm
            )
        )

        story.append(
            Paragraph(
                "<b>Interactive Conversation Experience</b>",
                body_style
            )
        )

        story.append(
            ReportLabImage(
                PRESENTATION_CHAT_2_IMAGE,
                width=165 * mm,
                height=72 * mm
            )
        )

        story.append(
            Spacer(
                1,
                5 * mm
            )
        )
    # --------------------------------------------------------
    # TECHNOLOGY
    # --------------------------------------------------------

    story.append(
        Paragraph(
            "10. Technology Behind the Project",
            heading_style
        )
    )

    for item in [
        "<b>Python</b> — core application development.",
        "<b>Streamlit</b> — interactive web application interface.",
        "<b>YouTubeTranscriptApi</b> — transcript retrieval.",
        "<b>LangChain</b> — document processing and RAG orchestration.",
        "<b>RecursiveCharacterTextSplitter</b> — transcript chunking.",
        "<b>OpenAI text-embedding-3-small</b> — semantic embeddings.",
        "<b>Chroma</b> — vector database and retrieval.",
        "<b>GPT-4o-mini</b> — AI answer generation.",
        "<b>ReportLab</b> — conversation and project PDF generation.",
        "<b>python-dotenv</b> — environment and API configuration."
    ]:
        story.append(
            Paragraph(
                "• " + item,
                bullet_style
            )
        )

    # --------------------------------------------------------
    # WORKFLOW / ARCHITECTURE
    # --------------------------------------------------------

    if os.path.exists(PRESENTATION_WORKFLOW_IMAGE):

        story.append(
            Spacer(
                1,
                3 * mm
            )
        )

        story.append(
            Paragraph(
                "<b>RAG Workflow & System Architecture</b>",
                body_style
            )
        )

        story.append(
            ReportLabImage(
                PRESENTATION_WORKFLOW_IMAGE,
                width=165 * mm,
                height=78 * mm
            )
        )

        story.append(
            Spacer(
                1,
                5 * mm
            )
        )


    # --------------------------------------------------------
    # CURRENT LIMITATIONS
    # --------------------------------------------------------

    story.append(
        Paragraph(
            "11. Current Limitations",
            heading_style
        )
    )

    for item in [
        "The YouTube video must have an accessible transcript.",
        "Answer quality depends on the quality and completeness of the transcript.",
        "The AI can only answer information supported by the retrieved transcript context.",
        "A valid OpenAI API configuration is required.",
        "Very long transcripts may require additional processing and optimization."
    ]:
        story.append(
            Paragraph(
                "• " + item,
                bullet_style
            )
        )


    # --------------------------------------------------------
    # FUTURE
    # --------------------------------------------------------

    story.append(
        Paragraph(
            "12. Future Possibilities",
            heading_style
        )
    )

    for item in [
        "Automatic video summarization.",
        "Automatic chapter and topic generation.",
        "Clickable transcript timestamps.",
        "Source snippets for each answer.",
        "Multilingual transcript support.",
        "Multilingual AI responses.",
        "Saved conversations.",
        "Video-specific FAQ generation.",
        "Automatic key-point extraction.",
        "Frequently asked question analytics.",
        "Improved transcript search and retrieval.",
        "More advanced video knowledge management."
    ]:
        story.append(
            Paragraph(
                "• " + item,
                bullet_style
            )
        )


    # --------------------------------------------------------
    # PROJECT VALUE
    # --------------------------------------------------------

    story.append(
        Paragraph(
            "13. What This Project Demonstrates",
            heading_style
        )
    )

    story.append(
        Paragraph(
            "This project demonstrates a complete practical "
            "Generative AI application rather than only an isolated "
            "LLM interaction. It combines data ingestion, transcript "
            "processing, document chunking, embeddings, vector "
            "storage, semantic retrieval, prompt engineering, "
            "LLM generation, conversational UI, session management, "
            "error handling and PDF generation into one working "
            "application.",
            body_style
        )
    )


    # --------------------------------------------------------
    # PROJECT SUMMARY
    # --------------------------------------------------------

    story.append(
        Paragraph(
            "14. Project Summary",
            heading_style
        )
    )

    story.append(
        Paragraph(
            "<b>YouTube AI Assistant makes long YouTube videos easier "
            "to understand and search.</b> Users provide a video link, "
            "ask questions in natural language, and receive answers "
            "based on the video's transcript. Behind the scenes, the "
            "application combines transcript processing, semantic "
            "search, vector retrieval and GPT-4o-mini through a RAG "
            "pipeline.",
            body_style
        )
    )


    # --------------------------------------------------------
    # PROJECT PITCH
    # --------------------------------------------------------

    story.append(
        Paragraph(
            "15. One-Line Project Pitch",
            heading_style
        )
    )

    story.append(
        Paragraph(
            "<b>Turn any accessible YouTube transcript into an "
            "interactive AI knowledge base that users can question "
            "in natural language.</b>",
            body_style
        )
    )


        
    # --------------------------------------------------------
    # BUILD PDF
    # --------------------------------------------------------

    document.build(story)
    buffer.seek(0)
    
    reader = PdfReader(buffer)
    writer = PdfWriter()
    total_pages = len(reader.pages)
    for page_number, page in enumerate(reader.pages):

        if page_number == total_pages - 1:
            page_width = float(A4[0])
            page_height = float(A4[1])
            contact_buffer = io.BytesIO()
            from reportlab.pdfgen import canvas
            contact_canvas = canvas.Canvas(
                contact_buffer,
                pagesize=A4
            )

            contact_canvas.setFillColor(
                colors.HexColor("#475569")
            )

            contact_canvas.setFont(
                "Helvetica-Bold",
                8
            )

            contact_canvas.drawRightString(
                page_width - 18 * mm,
                13 * mm,
                f"Developed by {DEVELOPER_NAME}"
            )

            contact_canvas.setFont(
                "Helvetica",
                7
            )

            contact_details = []

            if DEVELOPER_EMAIL:
                contact_details.append(
                    f"Email: {DEVELOPER_EMAIL}"
                )

            if DEVELOPER_PHONE:
                contact_details.append(
                    f"Phone: {DEVELOPER_PHONE}"
                )

            if contact_details:

                contact_canvas.drawRightString(
                    page_width - 18 * mm,
                    8 * mm,
                    "  |  ".join(contact_details)
                )

            contact_canvas.save()
            contact_buffer.seek(0)
            overlay = PdfReader(
                contact_buffer
            )

            page.merge_page(
                overlay.pages[0]
            )

        writer.add_page(page)

    final_buffer = io.BytesIO()

    writer.write(
        final_buffer
    )

    final_buffer.seek(0)

    return final_buffer.getvalue()    
   


# ============================================================
# PROJECT PRESENTATION PPTX
# ============================================================

def create_project_presentation_pptx():
    """
    Create a standard 16:9 PowerPoint presentation
    for users and clients.
    """

    presentation = Presentation()

    presentation.slide_width = Inches(13.333333)
    presentation.slide_height = Inches(7.5)

    # --------------------------------------------------------
    # HELPER: TITLE
    # --------------------------------------------------------

    def add_title(
        slide,
        title,
        subtitle=None
    ):

        title_box = slide.shapes.add_textbox(
            Inches(0.65),
            Inches(0.45),
            Inches(12.0),
            Inches(0.65)
        )

        text_frame = title_box.text_frame
        text_frame.text = title

        paragraph = text_frame.paragraphs[0]

        paragraph.font.size = Pt(28)
        paragraph.font.bold = True
        paragraph.font.name = "Aptos Display"


        if subtitle:

            subtitle_box = slide.shapes.add_textbox(
                Inches(0.68),
                Inches(1.12),
                Inches(11.8),
                Inches(0.55)
            )

            text_frame = subtitle_box.text_frame
            text_frame.text = subtitle

            paragraph = text_frame.paragraphs[0]

            paragraph.font.size = Pt(14)
            paragraph.font.name = "Aptos"


    # --------------------------------------------------------
    # HELPER: BULLETS
    # --------------------------------------------------------

    def add_bullets(
        slide,
        items,
        x=0.9,
        y=1.55,
        w=11.5,
        h=4.9
    ):

        box = slide.shapes.add_textbox(
            Inches(x),
            Inches(y),
            Inches(w),
            Inches(h)
        )

        text_frame = box.text_frame

        text_frame.word_wrap = True
        text_frame.clear()

        for index, item in enumerate(items):

            paragraph = (
                text_frame.paragraphs[0]
                if index == 0
                else text_frame.add_paragraph()
            )

            paragraph.text = item
            paragraph.font.size = Pt(20)
            paragraph.font.name = "Aptos"
            paragraph.space_after = Pt(10)


    # ========================================================
    # SLIDE 1 — COVER
    # ========================================================

    slide = presentation.slides.add_slide(
        presentation.slide_layouts[6]
    )

    add_title(
        slide,
        "YouTube AI Assistant",
        "Turn long YouTube videos into an interactive knowledge experience"
    )

    if os.path.exists(
        PRESENTATION_DASHBOARD_IMAGE
    ):

        slide.shapes.add_picture(
            PRESENTATION_DASHBOARD_IMAGE,
            Inches(0.65),
            Inches(1.75),
            width=Inches(12.0)
        )


    # ========================================================
    # SLIDE 2 — CHALLENGE
    # ========================================================

    slide = presentation.slides.add_slide(
        presentation.slide_layouts[6]
    )

    add_title(
        slide,
        "The Challenge",
        "Finding one answer inside a long video can take time"
    )

    add_bullets(
        slide,
        [
            "Long videos contain useful information but are difficult to search manually.",
            "Users often need one specific answer rather than the entire video.",
            "Traditional keyword search may not capture the meaning of a natural-language question.",
            "A conversational interface makes video knowledge easier to access."
        ]
    )


    # ========================================================
    # SLIDE 3 — SOLUTION
    # ========================================================

    slide = presentation.slides.add_slide(
        presentation.slide_layouts[6]
    )

    add_title(
        slide,
        "The Solution",
        "Ask questions naturally and receive transcript-grounded answers"
    )

    add_bullets(
        slide,
        [
            "Paste a YouTube video link.",
            "The application retrieves the available transcript.",
            "The transcript becomes a searchable knowledge base.",
            "Users ask questions in normal language.",
            "Relevant transcript context is retrieved before the AI generates an answer."
        ]
    )


    # ========================================================
    # SLIDE 4 — HOW IT WORKS
    # ========================================================

    slide = presentation.slides.add_slide(
        presentation.slide_layouts[6]
    )

    add_title(
        slide,
        "How It Works",
        "Retrieval-Augmented Generation pipeline"
    )

    add_bullets(
        slide,
        [
            "YouTube URL → Transcript",
            "Transcript → Text Chunks",
            "Text Chunks → OpenAI Embeddings",
            "Embeddings → Chroma Vector Database",
            "Retriever → Relevant Transcript Context",
            "GPT-4o-mini → Final Answer"
        ]
    )


    # ========================================================
    # SLIDE 5 — USER VALUE
    # ========================================================

    slide = presentation.slides.add_slide(
        presentation.slide_layouts[6]
    )

    add_title(
        slide,
        "What the User Gets",
        "A faster way to explore video-based knowledge"
    )

    add_bullets(
        slide,
        [
            "Natural-language video Q&A",
            "Faster access to specific information",
            "Transcript-grounded AI responses",
            "Conversation history",
            "PDF conversation export",
            "Simple and user-friendly interface"
        ]
    )


    # ========================================================
    # SLIDE 6 — USE CASES
    # ========================================================

    slide = presentation.slides.add_slide(
        presentation.slide_layouts[6]
    )

    add_title(
        slide,
        "Where It Can Be Used",
        "Examples of practical applications"
    )

    add_bullets(
        slide,
        [
            "Education and online learning",
            "Employee training and onboarding",
            "Technical tutorials and demonstrations",
            "Research and knowledge discovery",
            "Video-based documentation",
            "Personal learning"
        ]
    )


    # ========================================================
    # SLIDE 7 — TECHNOLOGY
    # ========================================================

    slide = presentation.slides.add_slide(
        presentation.slide_layouts[6]
    )

    add_title(
        slide,
        "Technology Behind the Experience",
        "The core technologies powering the application"
    )

    add_bullets(
        slide,
        [
            "Python — application development",
            "Streamlit — interactive user interface",
            "YouTubeTranscriptApi — transcript retrieval",
            "LangChain — RAG orchestration",
            "OpenAI Embeddings — semantic search",
            "Chroma — vector database",
            "GPT-4o-mini — answer generation",
            "ReportLab — PDF generation"
        ]
    )


    # ========================================================
    # SLIDE 8 — DASHBOARD SCREENSHOT
    # ========================================================

    slide = presentation.slides.add_slide(
        presentation.slide_layouts[6]
    )

    add_title(
        slide,
        "Product Experience",
        "YouTube AI Assistant dashboard"
    )

    if os.path.exists(
        PRESENTATION_DASHBOARD_IMAGE
    ):

        slide.shapes.add_picture(
            PRESENTATION_DASHBOARD_IMAGE,
            Inches(0.65),
            Inches(1.55),
            width=Inches(12.0)
        )


    # ========================================================
    # SLIDE 9 — RAG ARCHITECTURE
    # ========================================================

    slide = presentation.slides.add_slide(
        presentation.slide_layouts[6]
    )

    add_title(
        slide,
        "RAG Architecture",
        "From YouTube video to transcript-grounded AI answer"
    )

    if os.path.exists(
        PRESENTATION_WORKFLOW_IMAGE
    ):

        slide.shapes.add_picture(
            PRESENTATION_WORKFLOW_IMAGE,
            Inches(0.65),
            Inches(1.55),
            width=Inches(12.0)
        )


    # ========================================================
    # SLIDE 10 — FUTURE
    # ========================================================

    slide = presentation.slides.add_slide(
        presentation.slide_layouts[6]
    )

    add_title(
        slide,
        "Future Possibilities",
        "Potential extensions of the platform"
    )

    add_bullets(
        slide,
        [
            "Automatic video summaries",
            "Chapter and topic generation",
            "Timestamp-based source navigation",
            "Multilingual support",
            "Saved conversations and analytics",
            "Video-specific FAQ generation"
        ]
    )


    return_value = io.BytesIO()

    presentation.save(
        return_value
    )

    return_value.seek(0)

    return return_value.getvalue()

# ============================================================
# YOUTUBE URL / VIDEO ID HELPER
# ============================================================

def extract_youtube_video_id(value):
    """
    Extract a YouTube video ID from common YouTube URL formats.
    """

    if not value:
        return None

    value = value.strip()

    # --------------------------------------------------------
    # Direct Video ID support
    # --------------------------------------------------------

    if re.fullmatch(
        r"[A-Za-z0-9_-]{11}",
        value
    ):
        return value

    # --------------------------------------------------------
    # Add scheme if user pasted www.youtube.com/... without it
    # --------------------------------------------------------

    if not re.match(
        r"^https?://",
        value,
        re.IGNORECASE
    ):
        value = "https://" + value

    try:

        parsed = urlparse(value)

        host = (
            parsed.netloc
            .lower()
            .split(":")[0]
        )

        path_parts = [
            part
            for part in parsed.path.split("/")
            if part
        ]

        # ----------------------------------------------------
        # youtube.com/watch?v=VIDEO_ID
        # ----------------------------------------------------

        if host in {
            "youtube.com",
            "www.youtube.com",
            "m.youtube.com",
            "music.youtube.com"
        }:

            query = parse_qs(
                parsed.query
            )

            video_id = query.get(
                "v",
                [None]
            )[0]

            if video_id and re.fullmatch(
                r"[A-Za-z0-9_-]{11}",
                video_id
            ):
                return video_id

        # ----------------------------------------------------
        # youtu.be/VIDEO_ID
        # ----------------------------------------------------

        if host in {
            "youtu.be",
            "www.youtu.be"
        }:

            if path_parts:

                video_id = path_parts[0]

                if re.fullmatch(
                    r"[A-Za-z0-9_-]{11}",
                    video_id
                ):
                    return video_id

        # ----------------------------------------------------
        # youtube.com/shorts/VIDEO_ID
        # youtube.com/embed/VIDEO_ID
        # youtube.com/live/VIDEO_ID
        # youtube.com/v/VIDEO_ID
        # ----------------------------------------------------

        supported_paths = {
            "shorts",
            "embed",
            "live",
            "v"
        }

        if (
            host in {
                "youtube.com",
                "www.youtube.com",
                "m.youtube.com",
                "music.youtube.com"
            }
            and len(path_parts) >= 2
            and path_parts[0].lower()
            in supported_paths
        ):

            video_id = path_parts[1]

            if re.fullmatch(
                r"[A-Za-z0-9_-]{11}",
                video_id
            ):
                return video_id

    except Exception:
        return None

    return None

# ============================================================
# 6. SIDEBAR
# ============================================================
with st.sidebar:

    # ========================================================
    # YOUTUBE AI BRAND
    # ========================================================

    st.html("""
    <div class="sidebar-brand">

        <div class="sidebar-brand-icon">
            <span class="youtube-play-triangle">▶</span>
        </div>

        <div class="sidebar-brand-title">
            YouTube AI
        </div>

        <div class="sidebar-brand-text">
            AI assistant for intelligent
            video analysis
        </div>

    </div>
    """)


    # ========================================================
    # YOUTUBE VIDEO LINK
    # ========================================================

    st.html("""
    <div class="sidebar-section">
        🔗 &nbsp; YOUTUBE VIDEO LINK
    </div>
    """)

    youtube_url_input = st.text_input(
        "YouTube Video Link",
        value="",
        placeholder="Paste YouTube video link here...",
        label_visibility="collapsed"
    )

    st.caption(
        "Paste the YouTube video link here. "
        "I'll prepare the video automatically."
    )


    # ========================================================
    # PROCESS BUTTON
    # ========================================================

    process_button = st.button(
        "▶  Process Video",
        use_container_width=True
    )


    # ========================================================
    # PROJECT PRESENTATION
    # ========================================================

    project_presentation_pdf = (
        create_project_presentation_pdf()
    )

    st.download_button(
        label="📘 PROJECT DETAILS",
        data=project_presentation_pdf,
        file_name="YouTube_AI_Assistant_Project_Presentation.pdf",
        mime="application/pdf",
        type="primary",
        use_container_width=True,
        key="sidebar_project_presentation_download"
    )


    # ========================================================
    # DIVIDER
    # ========================================================

    st.html("""
    <div class="sidebar-divider"></div>
    """)


    # ========================================================
    # HOW IT WORKS
    # ========================================================

    st.html("""
    <div class="sidebar-section">
        ⚙ &nbsp; HOW IT WORKS
    </div>

    <div class="sidebar-list">

        <div>
            <span class="sidebar-number">1</span>
            Read the video transcript
        </div>

        <div>
            <span class="sidebar-number">2</span>
            Prepare the content
        </div>

        <div>
            <span class="sidebar-number">3</span>
            Understand the content
        </div>

        <div>
            <span class="sidebar-number">4</span>
            Build a searchable knowledge base
        </div>

        <div>
            <span class="sidebar-number">5</span>
            Find relevant information
        </div>

        <div>
            <span class="sidebar-number">6</span>
            Generate an AI answer
        </div>

    </div>
    """)


    # ========================================================
    # TECHNOLOGY
    # ========================================================

    st.html("""
    <div class="sidebar-section" style="margin-top:22px;">
        ◈ &nbsp; TECHNOLOGY
    </div>

    <div class="sidebar-list">

        <div>
            <span class="sidebar-bullet">•</span>
            YouTube Transcript API
        </div>

        <div>
            <span class="sidebar-bullet">•</span>
            LangChain
        </div>

        <div>
            <span class="sidebar-bullet">•</span>
            OpenAI Embeddings
        </div>

        <div>
            <span class="sidebar-bullet">•</span>
            Chroma Vector DB
        </div>

        <div>
            <span class="sidebar-bullet">•</span>
            GPT-4o-mini
        </div>

        <div>
            <span class="sidebar-bullet">•</span>
            Streamlit
        </div>

    </div>
    """)


    # ========================================================
    # SUPPORT / CONTACT
    # ========================================================

    st.html(f"""
    <div class="sidebar-support">

        <div class="sidebar-support-row">

            <div class="sidebar-support-icon">
                ✉
            </div>

            <div>

                <div class="sidebar-support-title">
                    Need Help?
                </div>

                <div class="sidebar-support-subtitle">
                    Feel free to contact me
                </div>

            </div>

        </div>

        <div class="sidebar-contact">

            {"<div>✉ &nbsp; " + DEVELOPER_EMAIL + "</div>"
                if DEVELOPER_EMAIL else ""}

            {"<div>♧ &nbsp; " + DEVELOPER_PHONE + "</div>"
                if DEVELOPER_PHONE else ""}

        </div>

    </div>

    <div class="sidebar-footer">

        Built and maintained by<br>

        <strong>
            {DEVELOPER_NAME}
        </strong>

        <span class="sidebar-heart">
            ♥
        </span>

    </div>
    """)
# ============================================================
# FULL PAGE YOUTUBE BACKGROUND IMAGE
# ============================================================

def get_background_image_base64(image_path):
    """Return a local image as Base64 for reliable CSS rendering."""
    if not os.path.exists(image_path):
        return ""

    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode("utf-8")


background_base64 = get_background_image_base64(
    YOUTUBE_THEME_BACKGROUND
)

# Use the local banner.png only.
# The user-provided/generated YouTube studio image is intentionally
# used across the complete Streamlit application background.
background_source = (
    f"url(data:image/png;base64,{background_base64})"
    if background_base64
    else "none"
)

st.html(
    f"""
    <style>

    html,
    body,
    .stApp,
    [data-testid="stAppViewContainer"] {{
        min-height: 100vh !important;
        background: transparent !important;
    }}

    .stApp {{
        background-image:
            linear-gradient(
                90deg,
                rgba(255, 255, 255, 0.07) 0%,
                rgba(255, 255, 255, 0.03) 45%,
                rgba(255, 240, 245, 0.05) 100%
            ),
            {background_source} !important;
        background-size: cover !important;
        background-position: center center !important;
        background-repeat: no-repeat !important;
        background-attachment: fixed !important;
    }}

    [data-testid="stAppViewContainer"],
    [data-testid="stMain"],
    [data-testid="stMainBlockContainer"],
    .main,
    .main .block-container,
    [data-testid="stHeader"],
    header {{
        background: transparent !important;
    }}

    [data-testid="stMain"] {{
        padding-top: 0 !important;
    }}

    [data-testid="stMainBlockContainer"] {{
        padding-top: 0 !important;
        margin-top: 0 !important;
    }}

    .main .block-container {{
        max-width: 1088px !important;
        width: 100% !important;
        padding-top: 0 !important;
        padding-bottom: 2rem !important;
        margin: 0 auto !important;
    }}

    .hero {{
        background: rgba(12, 22, 42, 0.34) !important;
        backdrop-filter: blur(2px);
        -webkit-backdrop-filter: blur(2px);
        border: 1px solid rgba(255, 255, 255, 0.24);
        border-radius: 20px;
        min-height: 300px;
        padding: 28px 54px 24px 54px;
        box-shadow: 0 20px 55px rgba(0, 0, 0, 0.18);
        margin-bottom: 25px !important;
    }}

    .empty-state {{

        background:
            rgba(10, 20, 38, 0.36) !important;

        backdrop-filter:
            blur(2px);

        -webkit-backdrop-filter:
            blur(2px);

        border:
            1px solid
            rgba(255, 255, 255, 0.24);

        border-radius:
            20px;

        min-height:
            304px;

        margin-top:
            0 !important;

        box-shadow:
            0 18px 45px
            rgba(0, 0, 0, 0.18);
    }}


        /* ============================================================
       FINAL CLEAN CARD VISIBILITY FIX
       ============================================================ */

    .premium-card {{
        background:
            linear-gradient(
                135deg,
                rgba(8, 18, 38, 0.88),
                rgba(13, 27, 55, 0.82)
            ) !important;

        backdrop-filter: blur(14px) !important;
        -webkit-backdrop-filter: blur(14px) !important;

        border:
            1px solid
            rgba(255, 255, 255, 0.22) !important;

        border-radius:
            20px !important;

        padding:
            24px 30px !important;

        box-shadow:
            0 18px 45px
            rgba(0, 0, 0, 0.30) !important;

        margin-bottom:
            14px !important;
    }}


    /* ============================================================
       CHAT SECTION TITLE
       ============================================================ */

    .section-title {{
        color:
            #ffffff !important;

        font-size:
            22px !important;

        font-weight:
            800 !important;

        line-height:
            1.3 !important;

        letter-spacing:
            -0.3px !important;

        margin-bottom:
            8px !important;

        text-shadow:
            0 2px 8px
            rgba(0, 0, 0, 0.55) !important;
    }}


    /* ============================================================
       CHAT SECTION DESCRIPTION
       ============================================================ */

    .section-subtitle {{
        color:
            #e8eef8 !important;

        font-size:
            14px !important;

        font-weight:
            500 !important;

        line-height:
            1.6 !important;

        max-width:
            950px !important;

        margin-bottom:
            0 !important;

        opacity:
            1 !important;

        text-shadow:
            0 2px 7px
            rgba(0, 0, 0, 0.60) !important;
    }}


    /* ============================================================
       STAT CARDS
       ============================================================ */

    .stat-card {{
        background:
            linear-gradient(
                135deg,
                rgba(8, 19, 40, 0.86),
                rgba(18, 31, 59, 0.80)
            ) !important;

        backdrop-filter:
            blur(12px) !important;

        -webkit-backdrop-filter:
            blur(12px) !important;

        border:
            1px solid
            rgba(255, 255, 255, 0.20) !important;

        border-radius:
            18px !important;

        box-shadow:
            0 16px 38px
            rgba(0, 0, 0, 0.28) !important;

        height:
            146px !important;

        min-height:
            146px !important;

        box-sizing:
            border-box !important;

        display:
            flex !important;

        flex-direction:
            column !important;

        justify-content:
            center !important;
    }}

    /* ============================================================
       STAT TEXT
       ============================================================ */

    .stat-label {{
        color:
            #bcd0ea !important;

        font-weight:
            800 !important;

        text-shadow:
            0 2px 6px
            rgba(0, 0, 0, 0.55) !important;
    }}


    .stat-value {{
        color:
            #ffffff !important;

        font-weight:
            800 !important;

        text-shadow:
            0 2px 8px
            rgba(0, 0, 0, 0.60) !important;
    }}


    /* ============================================================
       STATUS PILL
       ============================================================ */

    .status-pill {{
        background:
            rgba(5, 30, 25, 0.86) !important;

        border:
            1px solid
            rgba(74, 222, 128, 0.40) !important;

        color:
            #bbf7d0 !important;

        box-shadow:
            0 8px 24px
            rgba(0, 0, 0, 0.18) !important;
    }}


    /* ============================================================
       CONVERSATION — SINGLE LIGHT-BLACK PANEL
       ============================================================ */

    .st-key-conversation_history {{
        background: rgba(8, 15, 27, 0.68) !important;
        backdrop-filter: blur(14px) !important;
        -webkit-backdrop-filter: blur(14px) !important;
        border: 1px solid rgba(255, 255, 255, 0.14) !important;
        border-radius: 20px !important;
        box-shadow: 0 18px 45px rgba(0, 0, 0, 0.24) !important;
        overflow: hidden !important;
        margin-top: 0 !important;
        margin-bottom: 10px !important;
    }}

    .st-key-conversation_history p,
    .st-key-conversation_history li {{
        color: #f3f6fb !important;
        font-size: 14px !important;
        line-height: 1.65 !important;
    }}

    .st-key-conversation_history strong {{
        color: #ffffff !important;
        font-weight: 750 !important;
    }}

    .st-key-conversation_history hr {{
        border: 0 !important;
        border-top: 1px solid rgba(255, 255, 255, 0.07) !important;
        margin: 14px 0 !important;
    }}

    [data-testid="stChatMessage"] {{
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
    }}


    /* ============================================================
       CHAT INPUT
       ============================================================ */

    [data-testid="stChatInput"] {{
        background:
            rgba(7, 16, 32, 0.94) !important;

        border-radius:
            16px !important;

        box-shadow:
            0 10px 30px
            rgba(0, 0, 0, 0.35) !important;
    }}

    /* ============================================================
   PROJECT PRESENTATION
   ============================================================ */

    .presentation-wrapper {{
        background:
            linear-gradient(
                135deg,
                rgba(8, 18, 38, 0.92),
                rgba(18, 31, 59, 0.88)
            );

        backdrop-filter: blur(14px);
        -webkit-backdrop-filter: blur(14px);

        border: 1px solid rgba(255, 255, 255, 0.20);
        border-radius: 20px;

        padding: 28px 32px;
        margin-top: 18px;
        margin-bottom: 20px;

        box-shadow:
            0 18px 45px
            rgba(0, 0, 0, 0.30);
    }}

    .presentation-title {{
        color: #ffffff;
        font-size: 24px;
        font-weight: 800;
        margin-bottom: 8px;
    }}

    .presentation-subtitle {{
        color: #dbeafe;
        font-size: 14px;
        line-height: 1.65;
        margin-bottom: 20px;
    }}

    .presentation-section {{
        margin-top: 18px;
        padding: 18px 20px;
        border-radius: 16px;

        background: rgba(255, 255, 255, 0.055);
        border: 1px solid rgba(255, 255, 255, 0.10);
    }}

    .presentation-section-title {{
        color: #ffffff;
        font-size: 17px;
        font-weight: 800;
        margin-bottom: 8px;
    }}

    .presentation-section-text {{
        color: #dbe7f5;
        font-size: 13px;
        line-height: 1.7;
    }}

    .presentation-download {{
        margin-top: 20px;
        padding: 16px 18px;
        border-radius: 14px;

        background:
            linear-gradient(
                135deg,
                rgba(220, 38, 38, 0.18),
                rgba(37, 99, 235, 0.18)
            );

        border: 1px solid rgba(255, 255, 255, 0.16);
    }}


    /* ============================================================
    SCROLLABLE CONVERSATION CARD
    ============================================================ */

    .chat-conversation-card {{
        width: 100% !important;

        background:
            rgba(7, 12, 22, 0.64) !important;

        backdrop-filter:
            blur(12px) !important;

        -webkit-backdrop-filter:
            blur(12px) !important;

        border:
            1px solid
            rgba(255, 255, 255, 0.13) !important;

        border-radius:
            20px !important;

        padding:
            14px 16px !important;

        margin:
            16px 0 12px 0 !important;

        box-shadow:
            0 18px 45px
            rgba(0, 0, 0, 0.24) !important;

        box-sizing:
            border-box !important;
    }}


    /* ============================================================
    CHAT SCROLL AREA
    ============================================================ */

    .chat-scroll-area {{
        height:
            420px !important;

        display:
            flex !important;

        flex-direction:
            column-reverse !important;

        max-height:
            420px !important;

        overflow-y:
            auto !important;

        overflow-x:
            hidden !important;

        padding:
            2px 6px 2px 2px !important;

        scroll-behavior:
            smooth !important;

        overscroll-behavior:
            contain !important;

        scrollbar-width:
            thin;

        scrollbar-color:
            rgba(255,255,255,0.20)
            transparent;
    }}


    /* ============================================================
    CHAT SCROLLBAR
    ============================================================ */

    .chat-scroll-area::-webkit-scrollbar {{
        width:
            6px;
    }}

    .chat-scroll-area::-webkit-scrollbar-track {{
        background:
            transparent;
    }}

    .chat-scroll-area::-webkit-scrollbar-thumb {{
        background:
            rgba(255,255,255,0.20);

        border-radius:
            999px;
    }}

    .chat-scroll-area::-webkit-scrollbar-thumb:hover {{
        background:
            rgba(255,255,255,0.34);
    }}


    /* ============================================================
    CHAT MESSAGE ROWS
    ============================================================ */

    .chat-user-row,
    .chat-ai-row {{
        display:
            flex;

        align-items:
            flex-start;

        gap:
            12px;

        padding:
            13px 6px;

        box-sizing:
            border-box;
    }}

    .chat-ai-row {{
        border-top:
            1px solid
            rgba(255,255,255,0.08);
    }}


    /* ============================================================
    CHAT AVATAR
    ============================================================ */

    .chat-avatar {{
        width:
            36px;

        height:
            36px;

        min-width:
            36px;

        display:
            flex;

        align-items:
            center;

        justify-content:
            center;

        border-radius:
            11px;

        font-size:
            16px;
    }}

    .chat-user-avatar {{
        background:
            rgba(220, 38, 38, 0.18);

        border:
            1px solid
            rgba(248, 113, 113, 0.24);
    }}

    .chat-ai-avatar {{
        background:
            rgba(245, 158, 11, 0.16);

        border:
            1px solid
            rgba(251, 191, 36, 0.22);
    }}


    /* ============================================================
    CHAT TEXT
    ============================================================ */

    .chat-message-body {{
        min-width:
            0;

        flex:
            1;
    }}

    .chat-role {{
        color:
            rgba(255,255,255,0.72);

        font-size:
            12px;

        font-weight:
            700;

        margin-bottom:
            5px;
    }}

    .chat-content {{
        color:
            rgba(255,255,255,0.88);

        font-size:
            14px;

        line-height:
            1.65;

        font-weight:
            400;

        word-break:
            break-word;
    }}

    .chat-user-row,
    .chat-ai-row {{
        display: flex;
        align-items: flex-start;
        gap: 12px;
        padding: 12px 4px;
    }}

    .chat-ai-row {{
        border-top:
            1px solid
            rgba(255, 255, 255, 0.10);
    }}

    .chat-avatar {{
        width: 38px;
        height: 38px;
        min-width: 38px;

        display: flex;
        align-items: center;
        justify-content: center;

        border-radius: 12px;
        font-size: 18px;
    }}

    .chat-user-avatar {{
        background:
            linear-gradient(
                135deg,
                #ff4b4b,
                #ff7373
            );
    }}

    .chat-ai-avatar {{
        background:
            linear-gradient(
                135deg,
                #fbbf24,
                #f59e0b
            );
    }}

    .chat-role {{
        color: #cbd5e1;
        font-size: 12px;
        font-weight: 700;
        margin-bottom: 4px;
    }}

    .chat-content {{
        color: #f8fafc;
        font-size: 14px;
        line-height: 1.65;
    }}

    .chat-pdf-row {{
        display: flex;
        justify-content: flex-end;
        align-items: center;
        margin-top: 4px;
        margin-bottom: 8px;
    }}


    /* ============================================================
    GAP BEFORE ASK QUESTIONS CARD
    ============================================================ */

    .video-dashboard-gap {{
        height: 22px;
    }}

   

    .stVerticalBlock {{
        gap: 0 !important;
    }}

    </style>
    """
)

# ============================================================
# 7. HERO
# ============================================================
st.html("""
<div class="hero">

    <div class="hero-badge">
        ✦ AI-POWERED VIDEO INTELLIGENCE
    </div>


    <div class="hero-title">

        <span class="hero-gradient">
            YouTube
        </span>

        <span class="hero-main-title">
            AI Assistant
        </span>

    </div>


    <div class="hero-description">

        Turn long YouTube videos into an interactive
        knowledge base.

        <br>

        Ask questions naturally and get answers grounded
        in the video's transcript using
        Retrieval-Augmented Generation.

    </div>


    <!-- ====================================================
        HERO FEATURES
        ==================================================== -->

    <div class="hero-features">


        <!-- ASK ANYTHING -->

        <div class="hero-feature">

            <div class="hero-feature-icon">
                💬
            </div>

            <div>

                <div class="hero-feature-title">
                    Ask anything
                </div>

                <div class="hero-feature-text">
                    Get accurate answers
                </div>

            </div>

        </div>


        <!-- POWERED BY AI -->

        <div class="hero-feature">

            <div class="hero-feature-icon">
                ▤
            </div>

            <div>

                <div class="hero-feature-title">
                    Powered by AI
                </div>

                <div class="hero-feature-text">
                    Uses video transcript
                </div>

            </div>

        </div>


        <!-- SAVE TIME -->

        <div class="hero-feature">

            <div
                class="hero-feature-icon"
                style="
                    background:
                    rgba(126,20,92,0.55);
                "
            >
                ⚡
            </div>

            <div>

                <div class="hero-feature-title">
                    Save time
                </div>

                <div class="hero-feature-text">
                    Learn faster, smarter
                </div>

            </div>

        </div>


    </div>

</div>
""")

# ============================================================
# 8. API CONFIGURATION CHECK
# ============================================================

if not OPENAI_API_KEY:

    show_api_support_message(
        "missing"
    )


# ============================================================
# 9. PROCESS VIDEO
# ============================================================

if process_button:

    clean_video_id = extract_youtube_video_id(
        youtube_url_input
    )

    if not youtube_url_input.strip():

        st.html("""
        <div class="support-card">

            <div class="support-title">
                ⚠️ Please check the YouTube link
            </div>

            <div class="support-text">

                Please paste a valid YouTube video link.
                The video ID will be detected automatically.

            </div>

        </div>
        """)

    elif not OPENAI_API_KEY:

        st.session_state.api_error = True

        st.session_state.api_error_type = "missing"

        show_api_support_message(
            "missing"
        )

    else:

        # Reset previous error state
        st.session_state.api_error = False

        st.session_state.api_error_type = ""

        with st.spinner("Preparing your video..."):

            try:

                # ------------------------------------------------
                # STEP 1: FETCH TRANSCRIPT
                # ------------------------------------------------


                ytt_api = YouTubeTranscriptApi()

                try:

                    # First try the normal English-preferred transcript.
                    transcript = ytt_api.fetch(
                        clean_video_id,
                        languages=[
                            "en",
                            "en-US",
                            "en-GB"
                        ]
                    )

                except Exception:

                    # If English is unavailable, use any transcript
                    # available for the video.
                    transcript_list = ytt_api.list(
                        clean_video_id
                    )

                    available_transcripts = list(
                        transcript_list
                    )

                    if not available_transcripts:

                        raise ValueError(
                            "This video does not have an accessible transcript."
                        )

                    # Prefer the first available transcript.
                    selected_transcript = (
                        available_transcripts[0]
                    )

                    transcript = selected_transcript.fetch()

                full_text = " ".join(
                    item.text
                    for item in transcript
                )

                st.session_state.transcript_text = full_text

                if not full_text.strip():

                    raise ValueError(
                        "The transcript is empty."
                    )


                # ------------------------------------------------
                # STEP 2: CREATE DOCUMENT
                # ------------------------------------------------

                document = Document(
                    page_content=full_text,
                    metadata={
                        "video_id": clean_video_id
                    }
                )


                # ------------------------------------------------
                # STEP 3: SPLIT TEXT
                # ------------------------------------------------


                text_splitter = (
                    RecursiveCharacterTextSplitter(
                        chunk_size=1000,
                        chunk_overlap=200
                    )
                )

                docs = text_splitter.split_documents(
                    [document]
                )


                # ------------------------------------------------
                # STEP 4: EMBEDDINGS
                # ------------------------------------------------

                embeddings = OpenAIEmbeddings(
                    model="text-embedding-3-small",
                    api_key=OPENAI_API_KEY
                )


                # ------------------------------------------------
                # STEP 5: CHROMA
                # ------------------------------------------------


                persist_directory = (
                    f"./youtube_rag_db_{clean_video_id}"
                )

                vectorstore = Chroma.from_documents(
                    documents=docs,
                    embedding=embeddings,
                    persist_directory=persist_directory
                )


                # ------------------------------------------------
                # STEP 6: RETRIEVER
                # ------------------------------------------------

                retriever = vectorstore.as_retriever(
                    search_type="mmr",
                    search_kwargs={
                        "k": 8,
                        "fetch_k": 24
                    }
                )


                # ------------------------------------------------
                # STEP 7: PROMPT
                # ------------------------------------------------

                prompt = ChatPromptTemplate.from_template(
                    """
                You are a YouTube AI Assistant answering questions about the processed video.

                The retrieved transcript may have originally been written or spoken in any language.
                Understand the transcript regardless of its language, but ALWAYS answer the user in clear, natural English.

                Use ONLY the retrieved transcript context. Do not use outside knowledge and do not invent facts.

                Rules:

                1. Answer every question that can be supported by the transcript context.

                2. For broad questions such as "what is this video about",
                "summarize the video", "what are the key points", or
                "who is mentioned", synthesize the relevant context instead
                of refusing just because the question is broad.

                3. If the question is about a person, event, place, topic, or
                detail that is mentioned in the transcript, answer it directly
                and include the relevant details from the transcript.

                4. Do not claim that someone is visible in the video unless the
                transcript explicitly supports it.

                5. If the requested information is not supported by the transcript,
                reply politely:
                "I’m sorry, I don’t know based on the video transcript."

                6. If the question is unrelated to the video, reply politely:
                "I’m here to help with this video. I don’t have enough information
                in the video transcript to answer that."

                Retrieved transcript context:
                {context}

                User question:
                {question}

                Answer in English:
                """
                )

                # ------------------------------------------------
                # STEP 8: LLM
                # ------------------------------------------------


                llm = ChatOpenAI(
                    model="gpt-4o-mini",
                    temperature=0,
                    api_key=OPENAI_API_KEY
                )


                # ------------------------------------------------
                # STEP 9: RAG CHAIN
                # ------------------------------------------------

                rag_chain = (
                    {
                        "context": (
                            retriever
                            | format_docs
                        ),

                        "question": (
                            RunnablePassthrough()
                        )
                    }

                    | prompt
                    | llm
                    | StrOutputParser()
                )


                # ------------------------------------------------
                # SAVE SESSION
                # ------------------------------------------------

                st.session_state.rag_chain = (
                    rag_chain
                )

                st.session_state.video_id = (
                    clean_video_id
                )

                st.session_state.chunk_count = (
                    len(docs)
                )

                st.session_state.video_processed = (
                    True
                )

                st.session_state.chat_history = []

                st.session_state.api_error = (
                    False
                )

                st.session_state.api_error_type = (
                    ""
                )


                # ------------------------------------------------
                # SUCCESS
                # ------------------------------------------------

            except Exception as error:

                st.session_state.rag_chain = None

                st.session_state.video_processed = (
                    False
                )

                if is_api_error(error):

                    st.session_state.api_error = True

                    st.session_state.api_error_type = (
                        "invalid"
                    )

                    show_api_support_message(
                        "invalid"
                    )

                else:

                    st.html("""
                    <div class="support-card">

                        <div class="support-title">
                            ⚠️ We couldn't process this video
                        </div>

                        <div class="support-text">

                            We couldn't prepare this video
                            for questions right now.

                            <br><br>

                            Please check that the video has
                            an available transcript and try
                            again.

                            <br><br>

                            If the problem continues,
                            please contact the developer.

                        </div>

                    """.replace(
                        "</div>",
                        developer_contact_html()
                        + "</div>",
                        1
                    ))

                    st.session_state.api_error = True


# ============================================================
# 10. API SUPPORT MESSAGE
# ============================================================

if (
    st.session_state.api_error
    and not process_button
):

    show_api_support_message(
        st.session_state.api_error_type
        if st.session_state.api_error_type
        else "invalid"
    )


# ============================================================
# 11. VIDEO READY DASHBOARD
# ============================================================

if st.session_state.video_processed:

    st.html("""
    <div style="
        margin-bottom:15px;
    ">

        <span class="status-pill">

            <span class="status-dot"></span>

            VIDEO READY — YOU CAN START ASKING QUESTIONS

        </span>

    </div>
    """)


    col1, col2, col3 = st.columns(3)


    with col1:

        st.html(f"""
        <div class="stat-card">

            <div class="stat-icon">
                🎥
            </div>

            <div class="stat-label">
                Video
            </div>

            <div class="stat-value">
                Ready to explore
            </div>

        </div>
        """)

    with col2:

        st.html(f"""
        <div class="stat-card">

            <div class="stat-icon">
                🧩
            </div>

            <div class="stat-label">
                Content Sections
            </div>

            <div class="stat-value">
                {st.session_state.chunk_count:,}
            </div>

        </div>
        """)


    with col3:

        st.html("""
        <div class="stat-card">

            <div class="stat-icon">
                ✦
            </div>

            <div class="stat-label">
                AI Model
            </div>

            <div class="stat-value">
                GPT-4o-mini
            </div>

        </div>
        """)

   
# ============================================================
# 12. CHAT AREA
# ============================================================
if not st.session_state.video_processed:

    st.html("""
    <div class="empty-state">

        
        <div class="empty-title">
            Ready to explore your video?
        </div>

        <div class="empty-text">
            Paste a YouTube link in the sidebar and click
            <strong>Process Video</strong>.
            I'll turn the video's transcript into an
            AI-powered knowledge base you can explore naturally.
        </div>

        <div class="empty-quote">
            <span class="empty-quote-icon">💡</span>

            <span>
                Learn anything from any YouTube video,
                with the power of AI.
            </span>
        </div>

    </div>
    """)
    
else:

    # --------------------------------------------------------
    # GAP BETWEEN DASHBOARD CARDS AND ASK QUESTIONS CARD
    # --------------------------------------------------------

    st.html("""
    <div style="
        height: 15px;
        width: 100%;
    "></div>
    """)

    st.html("""
    <div class="premium-card">

        <div class="section-title">
            💬 Ask Questions About This Video
        </div>

        <div class="section-subtitle">

            Your video is ready. Ask anything about its content,
            key ideas, or important details — I'll answer using the video's transcript.

        </div>

    </div>
    """)


    # --------------------------------------------------------
    # CONVERSATION HISTORY — DISPLAY ONLY WHEN A QUESTION EXISTS
    # --------------------------------------------------------

    if st.session_state.chat_history:

        chat_rows = []

        for message in st.session_state.chat_history:

            if message["role"] == "user":
                role_label = "You"
                role_icon = "🔴"
                role_class = "chat-user-row"
                avatar_class = "chat-user-avatar"

            else:
                role_label = "AI Assistant"
                role_icon = "🤖"
                role_class = "chat-ai-row"
                avatar_class = "chat-ai-avatar"

            content = (
                str(message["content"])
                .strip()
                .replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\n", "<br>")
            )

            chat_rows.append(
                f"""
                <div class="{role_class}">

                    <div class="chat-avatar {avatar_class}">
                        {role_icon}
                    </div>

                    <div class="chat-message-body">

                        <div class="chat-role">
                            {role_label}
                        </div>

                        <div class="chat-content">
                            {content}
                        </div>

                    </div>

                </div>
                """
            )

        st.html(
            f"""
            <div class="chat-conversation-card">

                <div class="chat-scroll-area">
                    {"".join(reversed(chat_rows))}
                </div>

            </div>
            """
        )

        
    # --------------------------------------------------------
    # CHAT INPUT
    # --------------------------------------------------------

    query = st.chat_input(
        "Ask anything about this video..."
    )


    if query:

        st.session_state.chat_history.append(
            {
                "role": "user",
                "content": query
            }
        )

        with st.spinner(
            "Searching the video transcript..."
        ):

            try:

                # Handle greetings directly instead of searching the transcript.
                if is_greeting(query):
                    response = get_greeting_response()
                else:
                    response = (
                        st.session_state
                        .rag_chain
                        .invoke(query)
                    )

                st.session_state.chat_history.append(
                    {
                        "role": "assistant",
                        "content": response
                    }
                )

            except Exception as error:
                if is_api_error(error):
                    st.session_state.api_error = True
                    st.session_state.api_error_type = "invalid"
                    st.session_state.chat_history.append(
                        {
                            "role": "assistant",
                            "content": (
                                "I couldn't connect to the AI service "
                                "right now. Please check the OpenAI API "
                                "configuration and try again."
                            )
                        }
                    )

                else:

                    st.session_state.chat_history.append(
                        {
                            "role": "assistant",
                            "content": (
                                "I couldn't generate an answer "
                                "right now. Please try again."
                            )
                        }
                    )

        # Refresh once so the newly stored messages appear in the
        # single conversation panel above the chat input.
        st.rerun()


if st.session_state.chat_history:

    conversation_pdf = create_conversation_pdf(
        st.session_state.chat_history,
        st.session_state.video_id
    )

    st.download_button(
        label="📄 Save your Chat History",
        data=conversation_pdf,
        file_name=(
            f"video_conversation_"
            f"{st.session_state.video_id}.pdf"
        ),
        mime="application/pdf",
        help="Save this conversation as a PDF",
        key="conversation_pdf_download"
    )

# ============================================================
# 12. FOOTER
# ============================================================

st.html("""
<div class="youtube-footer">

    <div class="youtube-footer-links">

        WATCH
        <span>•</span>
        LEARN
        <span>•</span>
        EXPLORE
        <span>•</span>
        GROW
        <span>•</span>
        TOGETHER
    </div>
    <div class="youtube-footer-line"></div>
</div>
""")
