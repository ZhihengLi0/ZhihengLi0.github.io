#!/usr/bin/env python3
"""Build the CV as .docx (python-docx).  The PDF is made with LibreOffice:

    python3 cv/build_cv.py                 # -> cv/CV_ZhihengLi_<date>.docx (no phone number, for the website)
    python3 cv/build_cv.py --phone --out ~/cv_private/CV_ZhihengLi_<date>_full.docx
    module load libreoffice/12.05.22 && (cd ~ && libreoffice --headless --convert-to pdf --outdir <dir> <docx>)

Inline markup in the strings below: **bold**, *italic*.  A trailing "\t<date>" is right-aligned.
"""
import argparse, datetime, re
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Inches

FONT = "Times New Roman"
BODY = 10.5
PAGE_W, MARGIN_LR, MARGIN_TB = 8.5, 0.6, 0.5
RIGHT_TAB = Inches(PAGE_W - 2 * MARGIN_LR)


def add_runs(p, text, size=BODY):
    """**bold** and *italic* markup -> runs; '\t' stays a tab."""
    for tok in re.split(r"(\*\*.+?\*\*|\*.+?\*)", text):
        if not tok:
            continue
        bold = tok.startswith("**")
        ital = not bold and tok.startswith("*")
        r = p.add_run(tok.strip("*"))
        r.bold, r.italic = bold, ital
        r.font.name, r.font.size = FONT, Pt(size)
        r._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    return p


def para(doc, text="", size=BODY, align=None, before=0, after=0, left=0, hanging=0, right_tab=True):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before, pf.space_after, pf.line_spacing = Pt(before), Pt(after), 1.0
    if align:
        p.alignment = align
    if left or hanging:
        pf.left_indent = Inches(left)
        pf.first_line_indent = Inches(-hanging) if hanging else None
    if right_tab:
        pf.tab_stops.add_tab_stop(RIGHT_TAB, WD_TAB_ALIGNMENT.RIGHT)
    add_runs(p, text, size)
    return p


def rule(p):
    pPr = p._element.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    b = OxmlElement("w:bottom")
    for k, v in (("w:val", "single"), ("w:sz", "8"), ("w:space", "1"), ("w:color", "000000")):
        b.set(qn(k), v)
    bdr.append(b)
    pPr.append(bdr)


def section(doc, title):
    p = para(doc, f"**{title}**", size=BODY, before=6, after=2)
    rule(p)


def bullet(doc, text):
    """Hanging-indent bullet; a trailing '\t<date>' lands on the right margin."""
    return para(doc, "•\t" + text, left=0.25, hanging=0.17, after=1).paragraph_format


def entry(doc, left, right, bold=True, italic=False):
    t = f"**{left}**" if bold else (f"*{left}*" if italic else left)
    r = f"*{right}*" if italic and not bold else right
    para(doc, f"{t}\t{r}", before=2)


def build(phone):
    doc = Document()
    s = doc.sections[0]
    s.page_width, s.page_height = Inches(PAGE_W), Inches(11)
    s.left_margin = s.right_margin = Inches(MARGIN_LR)
    s.top_margin = s.bottom_margin = Inches(MARGIN_TB)
    st = doc.styles["Normal"]
    st.font.name, st.font.size = FONT, Pt(BODY)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    # bullet paragraphs: tab stop for the bullet glyph at the hanging position
    doc.styles["Normal"].paragraph_format.tab_stops.add_tab_stop(Inches(0.25))

    para(doc, "**Zhiheng (Kevin) Li**", size=16, align=WD_ALIGN_PARAGRAPH.CENTER, right_tab=False)
    contact = "li004628@umn.edu | " + ("+1 (612) 283-4843 | " if phone else "") + \
        "Minneapolis, MN | github.com/ZhihengLi0 | zhihengli0.github.io"
    para(doc, contact, size=10, align=WD_ALIGN_PARAGRAPH.CENTER, after=4, right_tab=False)

    # ---------------------------------------------------------------- education
    section(doc, "EDUCATION")
    entry(doc, "University of Minnesota, Twin Cities", "Minneapolis, MN")
    entry(doc, "Bachelor of Science in Mathematics  ·  Computer Science Minor", "Expected May 2028",
          bold=False, italic=True)
    bullet(doc, "**GPA:** 4.0/4.0 | **Honors:** Dean's List, Fall 2025, Spring 2026  ·  Tau Sigma Honor Society")
    bullet(doc, "**Relevant Coursework:** Calculus I, II & III, Linear Algebra & Differential Equations, Applied "
                "Linear Algebra, Proof-Based Sequences & Series, Physics for Scientists and Engineers I & II, "
                "Programming in Python and Java, Introduction to Algorithms & Data Structures, Introduction to "
                "Statistical Analysis, Cryptology & Number Theory")
    bullet(doc, "**Directed Reading Program (Fall 2026):** *Measure and Integral: An Introduction to Real Analysis* "
                "(Wheeden & Zygmund, 2nd ed.), as foundations for Optimal Transport, with mentor Anisha Jain, "
                "Ph.D. student in Mathematics, University of Minnesota")

    # ---------------------------------------------------------------- research
    section(doc, "RESEARCH EXPERIENCE")
    entry(doc, "SuperCDMS Dark Matter Experiment - School of Physics & Astronomy, University of Minnesota",
          "Minneapolis, MN")
    entry(doc, "SuperCDMS Researcher & Research Assistant to Prof. Yan Liu, UROP Fellow", "Oct 2025 - Present",
          bold=False, italic=True)
    bullet(doc, "Built **cdms_spotlight**, a multilingual natural-language query system for the SuperCDMS DQM "
                "database hosted at **SLAC**: server-side filtering scaling from ~1,300 to ~100,000 series, cuts on "
                "any of 10,929 ODB parameters per series, self-learning local NLP (TF-IDF + logistic regression); "
                "runs continuously as a Slack bot in the official collaboration workspace, pip-installable CLI; "
                "**presented at the SuperCDMS Software Meeting on July 14, 2026**; contributed to the follow-up "
                "DQM/ODB upgrade, prompted by cdms_spotlight, that moves non-metadata parameter cuts server-side so "
                "queries no longer pull data to the client.\tJul 2026")

    entry(doc, "SNOLAB Phonon Pulse Analysis (Run 4, Run 7)", "May 2026 - Present")
    bullet(doc, "Built the **phonon pulse template pipeline** for 13 SNOLAB R4 detectors (12 phonon channels each): "
                "event selection around the 10.37 keV activation line (~120 GB raw traces), free-pretrigger "
                "2-exponential fits with quality cuts; the 1×1 and PCA N×M ROOT templates serve as the R4 pulse "
                "templates for optimal-filter reprocessing; **presented at the SuperCDMS Analysis Working Group "
                "meeting on July 29, 2026**.\tMay–Jul 2026")
    bullet(doc, "Measured the **phonon collection efficiency** of detector Z7: converted every K-line pulse from "
                "current to power with the TES electrothermal equations and integrated the fitted pulse in closed "
                "form, validated against the official absorbed-energy RQ; **presented at the SuperCDMS Analysis "
                "Working Group meeting on September 23, 2026**; the current-to-power step uses the instant-sensor, "
                "infinite-feedback limit, refined below.\tJul–Sep 2026")
    bullet(doc, "**Full-model collection efficiency** of detector Z7: the efficiency above assumes an instant sensor "
                "with infinite feedback; derived the two coupled electro-thermal equations, linearised them to the "
                "complex impedance and fitted its four unknowns (current sensitivity, loop gain, thermal time "
                "constant, inductance) on Run 7 data with code written from the equations, reproducing the "
                "Irwin and Hilton (2005) two-time-constant response; the data show a **third pole, still an open "
                "question in the collaboration**; tested a second thermal body against the readout electronics, on "
                "the data and with the readout board measured alone on the bench, and placed it in the sensor; next, "
                "data from a second injection point, then a simulation to estimate the efficiency of low-energy "
                "events.\tOct 2026 - Present")
    bullet(doc, "Compared every **energy estimator** the Run 7 RQs allow on the same K-line events of Z7; a linear "
                "regression on the per-channel optimal-filter amplitudes gives several times the resolution of the "
                "plain OF sum.\tOct 2026")

    entry(doc, "Lab Instrumentation", "Oct 2025 - Jul 2026")
    bullet(doc, "Built **BlueFors-Alert**: minute-level PostgreSQL sync of the 64 GB CS2 database to a Raspberry Pi; "
                "Slack monitor with dual IDLE/COLD thresholds and time-range plots of any sensor; extended to a "
                "liquid-nitrogen scale.\tJul 2026")
    bullet(doc, "Extended RuOx (R31279 & R31839) **calibration curves** to 5 mK–300 K (PCHIP + Mott VRH); the "
                "generated .340 files are now in use on the lab's BlueFors dilution refrigerator.\tMay 2026")
    bullet(doc, "Built a **Python pipeline** for LED calibration data from 11 cryogenic Ge channels (5,200+ events, "
                "10+ GB raw): baseline subtraction, SNR-based glitch detection, peak alignment, exponential "
                "fitting.\tOct 2025")

    entry(doc, "Mouse Eye Tracking for Neural Alignment - Deisseroth & Mitra Labs, Stanford University", "Remote")
    entry(doc, "Research Collaborator (remote), with Kaiwen Sheng, Ph.D. student in Bioengineering",
          "Aug 2026 - Present", bold=False, italic=True)
    bullet(doc, "Set up eye keypoint tracking with **DeepLabCut** for infrared videos of mice recorded over multiple "
                "days (eight points: pupil ellipse, eyelids, eye corners); fixed the training recipe on the first "
                "video by controlled experiments with held-out test frames: batch size by validation loss, ResNet-50 "
                "among five backbones by test error and confidence, and the \"jump\" rule, best of three, for picking "
                "frames to correct; then defined every keypoint precisely for all later experiments.\tAug–Sep 2026")
    bullet(doc, "Ran **active learning on 19 videos from 3 mice**: fixed test and validation frames per video, first "
                "20 training frames by k-means on image appearance, then 20 frames per step where the jump rule "
                "finds unstable tracking, pre-labeled, corrected and retrained from scratch, until two steps improve "
                "the error by under 3 %. Most videos saturated at 20–40 labels; a new mouse needs its own labels "
                "(208 px to 10.4 px with 20) but starts better with every mouse seen; adding videos did not degrade "
                "earlier ones; 5 or 10 labels, checked with two subsets each, cannot replace 20. All data and "
                "scripts organised so that every step reproduces.\tSep–Oct 2026")
    bullet(doc, "Chose the **pupil-area rule** by comparing nine candidates against human ellipse labels on all test "
                "frames (four-point ellipse best, median filtering does not help) and built a per-frame **blink "
                "rule**: one eyelid-distance threshold per mouse set on frames judged by eye, pupil confidence as "
                "fallback; both run frame by frame, unsmoothed, for alignment with brain recordings.\tOct 2026")
    bullet(doc, "Built a self-correcting pipeline that **finds the frames the model gets wrong without new labels, "
                "repairs them from the network's own heatmaps under a geometric check learned from the human labels, "
                "and retrains on the result**; the retrained models make far fewer gross errors on frames labeled "
                "blind afterwards, and the chain runs on a new video without tuning.\tOct–Nov 2026")

    # ---------------------------------------------------------------- industry
    section(doc, "RESEARCH & INDUSTRY EXPERIENCE")
    entry(doc, "Low-Carbon Transition & Industrial Monitoring Platform", "Beijing / Shanghai, China")
    entry(doc, "Faculty-led Researcher (research phase) → Algorithm Engineer (industry phase), Shanghai "
               "Huaruizhongxin Technology Co., Ltd.", "", bold=False, italic=True)
    para(doc, "Oct 2024 - May 2025; Aug 2026 - Sep 2026", align=WD_ALIGN_PARAGRAPH.RIGHT, right_tab=False)
    bullet(doc, "Contributed to a faculty-led low-carbon transition project for China's steel industry, supporting a "
                "proposal and implementation plan for national ministry consideration through data collection, "
                "database construction, and quantitative analysis.")
    bullet(doc, "During the industry implementation phase, contributed to an energy and carbon-emissions monitoring "
                "platform by designing an algorithm prototype for industrial time-series data quality governance, "
                "translating meter/point metadata and quality flags into rules for cumulative, instantaneous, "
                "consumption, enumeration, and statistical measurements.")
    bullet(doc, "Implemented historical and real-time cleaning for range, constant, trend, missing-data, and "
                "communication anomalies; used same-time historical profiles for outage imputation and "
                "cumulative-to-consumption attribution; validated multi-day data with plots and regression tests, "
                "then packaged and delivered the work.")

    # ---------------------------------------------------------------- projects
    section(doc, "SELECTED PROJECTS")
    entry(doc, "College Students Entrepreneurial Project & Entrepreneurship Roadshow", "Beijing, China")
    entry(doc, "Team Leader", "Oct 2024 - Jun 2025", bold=False, italic=True)
    bullet(doc, "Led a student team to design a **billiards rest** that combines the high and low bridge heads into one;")
    bullet(doc, "Took it from CAD model to 3D-printed prototypes, iterating through player testing and market "
                "research; received **national project approval with seed funding**; **Chinese utility patent** "
                "under CNIPA review.")
    bullet(doc, "Invited to present at the **Tsinghua X-Lab & Kering Entrepreneurship Roadshow** to industry leaders.")

    # ---------------------------------------------------------------- awards
    section(doc, "AWARDS & HONORS")
    for t in ("**UROP** (Undergraduate Research Opportunities Program) Fellowship, University of Minnesota, Summer 2026",
              "**1st Prize**, Beijing University Student Innovation & Entrepreneurship Competition (Science & Technology Track)",
              "**1st Prize**, National College E-Commerce “Innovation, Creativity & Entrepreneurship” Challenge, "
              "Campus Level | Team Leader",
              "**3rd Prize**, China International College Students’ Innovation Competition, Campus Level (2025) | Team Leader",
              "**2nd Prize**, NECCS National English Competition for College Students"):
        bullet(doc, t)

    # ---------------------------------------------------------------- activities
    section(doc, "ACTIVITIES & SERVICE")
    for t in ("**Special Olympics USA Games Minnesota 2026**, Tennis Division - Volunteer Ball Retriever | June 25, 2026",
              "**Chinese Student Union**, UMN - Active Member | Oct 2025 - Present",
              "**Chinese Students & Scholars Association**, UMN - Volunteer Calculus I Tutor | 2025 - Present",
              "**College Table Tennis Club / Football Club / Guitar Club** - Team Member | 2023 - 2024"):
        bullet(doc, t)

    # ---------------------------------------------------------------- skills
    section(doc, "PROFESSIONAL SKILLS")
    bullet(doc, "**Techniques:** Cryogenic instrumentation, signal processing, data pipeline development, thermometer "
                "calibration curve fitting, real-time PostgreSQL sync, local NLP, Slack bot development, waveform "
                "template construction, TES / SQUID readout modelling, electronics transfer-function measurement, "
                "MIDAS DAQ control, keypoint tracking, active learning")
    bullet(doc, "**Tools:** Python (NumPy, Pandas, SciPy, Matplotlib), C++, scikit-learn, DeepLabCut, PyTorch, "
                "PostgreSQL, Java, ROOT, MIDAS / ODB, Git, Linux, Slurm, MSI HPC (Agate)")
    bullet(doc, "**Languages:** Mandarin Chinese (Native), English (Fluent; IELTS 7.0)")
    return doc


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--phone", action="store_true", help="include the phone number")
    ap.add_argument("--out", default=None)
    a = ap.parse_args()
    out = a.out or f"cv/CV_ZhihengLi_{datetime.date.today():%Y%m%d}.docx"
    build(a.phone).save(out)
    print(out)
