import os
from io import BytesIO
from pathlib import Path
from typing import Dict, List

import pandas as pd
import streamlit as st
from sqlalchemy import create_engine, text

st.set_page_config(page_title="FitCheck Query Portal", layout="wide")

DB_URL = os.getenv("POSTGRES_URL", "postgresql+psycopg2://postgres:postgresDARA1@localhost:5432/fitcheck_main")

# Path to the cleaned export workbook produced by the ETL pipeline
_APP_DIR = Path(__file__).resolve().parent
CLEANED_EXPORT_PATH = Path(os.getenv("FITCHECK_EXPORT_PATH", str(_APP_DIR.parent / "FitCheck_os" / "fitcheck_cleaned_export.xlsx")))

TABLE_CONFIG: Dict[str, Dict[str, List[str]]] = {
    "participants": {
        "default_columns": ["participant_id"],
        "filters": ["participant_id"],
    },
    "demographics_health": {
        "default_columns": [
            "participant_id", "sex_assigned_at_birth", "ethnicity", "age_years",
            "daily_physical_activity", "recurring_chronic_pain", "general_health_rating",
            "ergonomic_familiarity", "weekly_desk_hours_time",
        ],
        "filters": [
            "participant_id", "sex_assigned_at_birth", "ethnicity", "age_years",
            "recurring_chronic_pain", "general_health_rating", "ergonomic_familiarity",
        ],
    },
    "anthropometry": {
        "default_columns": [
            "participant_id", "session_date", "body_mass_kg", "bmi_kg_m2",
            "hip_breadth_sitting_cm", "elbow_height_sitting_cm",
            "buttock_popliteal_length_cm", "popliteal_height_sitting_cm",
        ],
        "filters": [
            "participant_id", "session_date", "body_mass_kg", "bmi_kg_m2",
            "body_fat_percent", "skeletal_muscle_percent",
        ],
    },
    "chair_selection_pre": {
        "default_columns": [
            "participant_id", "pre_selected_chair_code", "pre_selection_basis",
            "pre_selection_confidence_3_3",
            "pre_selection_basis_1", "pre_selection_basis_2", "pre_selection_basis_3",
        ],
        "filters": [
            "participant_id", "pre_selected_chair_code", "pre_selection_confidence_3_3",
            "pre_chair_familiarity_check_1", "pre_chair_familiarity_check_2",
        ],
    },
    "chair_selection_post": {
        "default_columns": [
            "participant_id", "post_selected_chair_code", "post_selection_basis",
            "post_selection_confidence_3_3", "selection_changed", "selection_change_reason",
            "post_selection_basis_1", "post_selection_basis_2",
        ],
        "filters": [
            "participant_id", "post_selected_chair_code", "post_selection_confidence_3_3",
            "selection_changed",
        ],
    },
    "chair_lookup": {
        "default_columns": ["chair_code", "chair_id", "chair_model", "chair_variant", "chair_name_display"],
        "filters": ["chair_code", "chair_id", "chair_model", "chair_variant"],
    },
    "chair_feature_bounds_wide": {
        "default_columns": [
            "chair_code", "chair_id", "chair_model", "chair_variant", "chair_name_display",
            "seat_height_min_cm", "seat_height_max_cm", "seat_depth_min_cm", "seat_depth_max_cm",
            "armrest_height_min_cm", "armrest_height_max_cm", "lumbar_min_unit", "lumbar_max_unit",
        ],
        "filters": ["chair_code", "chair_id", "chair_model", "chair_variant"],
    },
    "chair_trials": {
        "default_columns": [
            "participant_id", "session_date", "chair_code", "trial_order",
            "overall_comfort_3_3", "discomfort_seat_3_3", "discomfort_back_3_3",
            "expected_vs_experienced_overall_3_3", "prolonged_willingness_3_3",
            "geometric_fit_1", "geometric_fit_2", "geometric_fit_3",
            "headrest_dependency", "interfere",
        ],
        "filters": [
            "participant_id", "session_date", "chair_code", "trial_order",
            "overall_comfort_3_3", "prolonged_willingness_3_3",
            "headrest_dependency", "interfere",
        ],
    },
}

# ---------------------------------------------------------------------------
# CODEBOOK  — variable descriptions and coded value labels from the PDF docs
# ---------------------------------------------------------------------------
CODEBOOK: Dict[str, Dict] = {
    # ── Demographics & Health ────────────────────────────────────────────────
    "participant_id": {
        "question": "Participant ID",
        "type": "Text",
        "options": None,
    },
    "sex_assigned_at_birth": {
        "question": "Sex Assigned at Birth",
        "type": "Categorical",
        "options": {1: "Male", 2: "Female", 3: "Non-binary", 4: "Prefer not to say"},
    },
    "ethnicity": {
        "question": "Ethnicity",
        "type": "Categorical",
        "options": {
            1: "American Indian or Alaska Native",
            2: "Asian",
            3: "Black or African American",
            4: "Hispanic/Latino/Spanish",
            5: "Middle Eastern or North African",
            6: "Native Hawaiian or Pacific Islander",
            7: "White",
            8: "Other",
        },
    },
    "birthplace_city_state": {
        "question": "Birthplace (city, state, country)",
        "type": "Text",
        "options": None,
    },
    "age_years": {
        "question": "Age",
        "type": "Numeric",
        "options": None,
    },
    "daily_physical_activity": {
        "question": "Daily Physical Activity level",
        "type": "Ordinal",
        "options": {
            1: "Mostly sedentary (≤1 h/wk)",
            2: "Light activity (2–4 h/wk)",
            3: "Moderate activity (4–6 h/wk)",
            4: "Vigorous activity (≥6 h/wk)",
        },
    },
    "weekly_desk_hours_time": {
        "question": "Hours per week spent sitting at a desk",
        "type": "Ordinal",
        "options": {1: "Less than 10", 2: "10–20", 3: "21–40", 4: "More than 40"},
    },
    "general_health_rating": {
        "question": "Overall health rating",
        "type": "Ordinal",
        "options": {1: "Poor", 2: "Fair", 3: "Good", 4: "Very good", 5: "Excellent"},
    },
    "work_before_break_time": {
        "question": "Work duration before a break of at least 5 minutes",
        "type": "Ordinal",
        "options": {
            1: "30 min or less",
            2: "31–60 min",
            3: "61–90 min",
            4: "91–120 min",
            5: "More than 120 min",
        },
    },
    "ergonomic_familiarity": {
        "question": "Familiarity with ergonomic seating principles",
        "type": "Ordinal",
        "options": {
            1: "Not at all familiar",
            2: "Slightly familiar",
            3: "Moderately familiar",
            4: "Very familiar",
            5: "Extremely familiar",
        },
    },
    "recurring_chronic_pain": {
        "question": "Do you have recurring chronic pain?",
        "type": "Binary",
        "options": {1: "Yes", 0: "No"},
    },
    "pain_location": {
        "question": "If yes, where is the pain located?",
        "type": "Text",
        "options": None,
    },
    "posture_most_used": {
        "question": "Which posture do you most frequently use during desk work?",
        "type": "Categorical",
        "options": {
            1: "Upright", 2: "Perched", 3: "Recline", 4: "Leg(s) propped",
            5: "Lotus", 6: "Feet Up to the Side", 7: "Leg over Foot", 8: "Legs Crossed",
            9: "Knee Up", 10: "Both Knees Up", 11: "Side Sitting", 12: "Kneeling",
        },
    },
    "posture_switch_to": {
        "question": "Which other posture do you use occasionally?",
        "type": "Categorical",
        "options": {
            1: "Upright", 2: "Perched", 3: "Recline", 4: "Leg(s) propped",
            5: "Lotus", 6: "Feet Up to the Side", 7: "Leg over Foot", 8: "Legs Crossed",
            9: "Knee Up", 10: "Both Knees Up", 11: "Side Sitting", 12: "Kneeling",
        },
    },
    "posture_change_frequency": {
        "question": "During a typical 1-hour session, how often do posture changes occur?",
        "type": "Ordinal",
        "options": {
            1: "Rarely (0–1 times)",
            2: "Occasionally (2–3 times)",
            3: "Frequently (4–6 times)",
            4: "Very frequently (7+ times)",
        },
    },
    "seat_continuously_change_freq": {
        "question": "How long do you sit continuously before changing posture or standing?",
        "type": "Ordinal",
        "options": {
            1: "30 min or less",
            2: "31–60 min",
            3: "61–90 min",
            4: "91–120 min",
            5: "More than 120 min",
        },
    },
    **{f"posture_list_familarity_{i}": {
        "question": f"Frequency of using posture {i}",
        "type": "Ordinal",
        "options": {1: "Never", 2: "Sometimes", 3: "Often"},
    } for i in range(1, 13)},
    **{f"hypermobility_matrix_{i}": {
        "question": [
            "Can you place your hands flat on the floor without bending your knees?",
            "Can you bend your thumb to touch your forearm?",
            "As a child, did you amuse others by contorting your body / doing the splits?",
            "As a child/teenager, did your kneecap or shoulder dislocate on more than one occasion?",
            "Do you consider yourself double-jointed?",
        ][i - 1],
        "type": "Binary",
        "options": {1: "Yes", 0: "No"},
    } for i in range(1, 6)},
    # ── Pre-Selection ────────────────────────────────────────────────────────
    **{f"pre_chair_familiarity_check_{i}": {
        "question": f"Familiarity with chair C{i:02d}",
        "type": "Ordinal",
        "options": {
            1: "Never seen before",
            2: "Seen but never used",
            3: "Used briefly",
            4: "Used regularly",
        },
    } for i in range(1, 12)},
    "pre_selected_chair_code": {
        "question": "Initial chair choice (before trials)",
        "type": "Categorical",
        "options": {i: f"C{i:02d}" for i in range(1, 12)},
    },
    **{f"pre_selection_basis_{i}": {
        "question": "Reason for initial chair selection — " + [
            "It looks the most comfortable",
            "It appears to provide the best back support",
            "The seat cushion looks comfortable",
            "The chair design looks ergonomic",
            "The chair looks adjustable or flexible",
            "The chair appears to fit my body posture",
            "The chair looks high quality or well-built",
            "I like the visual design or aesthetics",
            "I have sat in it before today",
            "Other (please specify)",
        ][i - 1],
        "type": "Binary",
        "options": {1: "Selected", 0: "Not selected"},
    } for i in range(1, 11)},
    "pre_selection_basis": {
        "question": "Other basis for initial selection (open text)",
        "type": "Text",
        "options": None,
    },
    "pre_selection_confidence_3_3": {
        "question": "Confidence in initial chair selection",
        "type": "Ordinal",
        "options": {
            -3: "Very unconfident", -2: "Moderately unconfident", -1: "Slightly unconfident",
            0: "Neutral",
            1: "Slightly confident", 2: "Moderately confident", 3: "Very confident",
        },
    },
    # ── Post-Selection ───────────────────────────────────────────────────────
    "post_selected_chair_code": {
        "question": "Best overall chair after testing all chairs",
        "type": "Categorical",
        "options": {i: f"C{i:02d}" for i in range(1, 12)},
    },
    **{f"post_selection_basis_{i}": {
        "question": "Reason for final chair selection — " + [
            "It looks the most comfortable",
            "It appears to provide the best back support",
            "The seat cushion looks comfortable",
            "The chair design looks ergonomic",
            "The chair looks adjustable or flexible",
            "The chair appears to fit my body posture",
            "The chair looks high quality or well-built",
            "I like the visual design or aesthetics",
            "I have sat in it before today",
            "Other (please specify)",
        ][i - 1],
        "type": "Binary",
        "options": {1: "Selected", 0: "Not selected"},
    } for i in range(1, 11)},
    "post_selection_basis": {
        "question": "Other basis for final selection (open text)",
        "type": "Text",
        "options": None,
    },
    "post_selection_confidence_3_3": {
        "question": "Confidence in final chair selection",
        "type": "Ordinal",
        "options": {
            -3: "Very unconfident", -2: "Moderately unconfident", -1: "Slightly unconfident",
            0: "Neutral",
            1: "Slightly confident", 2: "Moderately confident", 3: "Very confident",
        },
    },
    "selection_changed": {
        "question": "Did your preferred chair change from pre- to post-selection?",
        "type": "Binary",
        "options": {0: "No", 1: "Yes"},
    },
    "selection_change_reason": {
        "question": "What most influenced the change in preference?",
        "type": "Categorical",
        "options": {
            1: "Actual sitting comfort",
            2: "Better fit after adjustment",
            3: "Back support",
            4: "Seat comfort over time",
            5: "Pressure/discomfort",
            6: "Ease of adjustment",
            7: "Other",
        },
    },
    # ── Chair Trials — metadata ───────────────────────────────────────────────
    "headrest_dependency": {
        "question": "Chair has a headrest",
        "type": "Binary",
        "options": {1: "True", 0: "False"},
    },
    "headrest_dependency_user": {
        "question": "Did you use the headrest?",
        "type": "Binary",
        "options": {1: "Yes", 0: "No"},
    },
    "interfere": {
        "question": "Do any features make it harder to sit comfortably, maintain posture, or move naturally?",
        "type": "Binary",
        "options": {1: "Yes", 0: "No"},
    },
    "interfere_location_selected_choice": {
        "question": "Interference location / features",
        "type": "Categorical",
        "options": {
            1: "Seat edge / frame",
            2: "Armrest Support",
            3: "Recline Motion",
            4: "Lumbar Area",
            5: "Headrest",
            6: "Others",
        },
    },
    "interfere_severity": {
        "question": "Severity of physical interference",
        "type": "Ordinal",
        "options": {1: "Mild", 2: "Moderate", 3: "Severe"},
    },
    # ── Chair Trials — fit scales (-3 to +3) ─────────────────────────────────
    "elbow_fit": {
        "question": "Armrest height after adjustment",
        "type": "Ordinal",
        "options": {
            -3: "Much too low", -2: "Moderately too low", -1: "Slightly too low",
            0: "About right",
            1: "Slightly too high", 2: "Moderately too high", 3: "Much too high",
        },
    },
    "seat_depth_fit": {
        "question": "Seat depth after adjustment",
        "type": "Ordinal",
        "options": {
            -3: "Much too shallow", -2: "Moderately too shallow", -1: "Slightly too shallow",
            0: "About right",
            1: "Slightly too deep", 2: "Moderately too deep", 3: "Much too deep",
        },
    },
    "seat_width_fit": {
        "question": "Seat width",
        "type": "Ordinal",
        "options": {
            -3: "Much too narrow", -2: "Moderately too narrow", -1: "Slightly too narrow",
            0: "About right",
            1: "Slightly too wide", 2: "Moderately too wide", 3: "Much too wide",
        },
    },
    "seat_height_fit": {
        "question": "Seat height after adjustment",
        "type": "Ordinal",
        "options": {
            -3: "Much too low", -2: "Moderately too low", -1: "Slightly too low",
            0: "About right",
            1: "Slightly too high", 2: "Moderately too high", 3: "Much too high",
        },
    },
    "backrest_width_fit": {
        "question": "Backrest width",
        "type": "Ordinal",
        "options": {
            -3: "Much too narrow", -2: "Moderately too narrow", -1: "Slightly too narrow",
            0: "About right",
            1: "Slightly too wide", 2: "Moderately too wide", 3: "Much too wide",
        },
    },
    "backrest_height_fit": {
        "question": "Backrest height",
        "type": "Ordinal",
        "options": {
            -3: "Much too low", -2: "Moderately too low", -1: "Slightly too low",
            0: "About right",
            1: "Slightly too high", 2: "Moderately too high", 3: "Much too high",
        },
    },
    "headrest_height_fit": {
        "question": "Headrest height",
        "type": "Ordinal",
        "options": {
            -3: "Much too low", -2: "Moderately too low", -1: "Slightly too low",
            0: "About right",
            1: "Slightly too high", 2: "Moderately too high", 3: "Much too high",
        },
    },
    "made_for_me_fit": {
        "question": "Do you feel this chair is made for you?",
        "type": "Ordinal",
        "options": {
            -3: "Strongly disagree", -2: "Moderately disagree", -1: "Slightly disagree",
            0: "Neutral",
            1: "Slightly agree", 2: "Moderately agree", 3: "Strongly agree",
        },
    },
    **{f"geometric_fit_{i}": {
        "question": [
            "The armrests provide good support",
            "The seat feels comfortable",
            "The chair feels comfortable when I first sit down",
            "The chair seat height fits me",
            "Backrest width fits you",
            "Backrest height fits you",
            "Headrest height fits you",
        ][i - 1],
        "type": "Ordinal",
        "options": {
            -3: "Strongly disagree", -2: "Moderately disagree", -1: "Slightly disagree",
            0: "Neutral",
            1: "Slightly agree", 2: "Moderately agree", 3: "Strongly agree",
        },
    } for i in range(1, 8)},
    # ── Chair Trials — comfort scales ─────────────────────────────────────────
    "discomfort_seat_3_3": {
        "question": "Seat comfort / discomfort rating",
        "type": "Ordinal",
        "options": {
            -3: "Extremely uncomfortable", -2: "Very uncomfortable", -1: "Slightly uncomfortable",
            0: "Neutral",
            1: "Slightly comfortable", 2: "Very comfortable", 3: "Extremely comfortable",
        },
    },
    "discomfort_back_3_3": {
        "question": "Back support comfort / discomfort rating",
        "type": "Ordinal",
        "options": {
            -3: "Extremely uncomfortable", -2: "Very uncomfortable", -1: "Slightly uncomfortable",
            0: "Neutral",
            1: "Slightly comfortable", 2: "Very comfortable", 3: "Extremely comfortable",
        },
    },
    "overall_comfort_3_3": {
        "question": "Overall comfort / discomfort rating",
        "type": "Ordinal",
        "options": {
            -3: "Extremely uncomfortable", -2: "Very uncomfortable", -1: "Slightly uncomfortable",
            0: "Neutral",
            1: "Slightly comfortable", 2: "Very comfortable", 3: "Extremely comfortable",
        },
    },
    "chair_aesthetics_1": {
        "question": "How attractive do you find this chair?",
        "type": "Ordinal",
        "options": {-3: "Very unattractive", 0: "Average", 3: "Very attractive"},
    },
    "chair_aesthetics_2": {
        "question": "Overall, how much do you like this chair?",
        "type": "Ordinal",
        "options": {-3: "Dislike very much", 0: "Neutral", 3: "Like very much"},
    },
    "chair_aesthetics_30": {
        "question": "Overall, how comfortable do you find this chair?",
        "type": "Ordinal",
        "options": {-3: "Extremely uncomfortable", 0: "Neither comfortable nor uncomfortable", 3: "Extremely comfortable"},
    },
    "expected_vs_experienced_seat_3_3": {
        "question": "Compared to initial impression — seat comfort",
        "type": "Ordinal",
        "options": {-3: "Much worse than expected", 0: "About as expected", 3: "Much better than expected"},
    },
    "expected_vs_experienced_back_3_3": {
        "question": "Compared to initial impression — back support",
        "type": "Ordinal",
        "options": {-3: "Much worse than expected", 0: "About as expected", 3: "Much better than expected"},
    },
    "expected_vs_experienced_overall_3_3": {
        "question": "Compared to initial impression — overall comfort",
        "type": "Ordinal",
        "options": {-3: "Much worse than expected", 0: "About as expected", 3: "Much better than expected"},
    },
    "prolonged_willingness_3_3": {
        "question": "Would you be willing to use this chair for extended periods?",
        "type": "Categorical",
        "options": {0: "No", 1: "Yes", 2: "Unsure"},
    },
    # ── Chair Trials — improvement suggestions ────────────────────────────────
    **{f"improve_suggestion_{i}": {
        "question": "What could be improved about this chair? — " + [
            "Fit issue",
            "Physical interference",
            "Lack of support",
            "Adjustment limitation",
            "Material comfort issue",
            "Other",
            "None",
        ][i - 1],
        "type": "Binary",
        "options": {1: "Selected", 0: "Not selected"},
    } for i in range(1, 8)},
    **{f"improve_suggestion_loc_{i}": {
        "question": "Control/mechanism improvement — " + [
            "Ease of adjusting controls",
            "Reachability of controls",
            "Recline smoothness",
            "Stability during recline",
            "Other",
            "None",
        ][i - 1],
        "type": "Binary",
        "options": {1: "Selected", 0: "Not selected"},
    } for i in range(1, 7)},
}

# Which variables belong to which table (for codebook filtering)
_CODEBOOK_TABLE_MAP: Dict[str, List[str]] = {
    "demographics_health": [
        "participant_id", "sex_assigned_at_birth", "ethnicity", "birthplace_city_state",
        "age_years", "daily_physical_activity", "weekly_desk_hours_time", "general_health_rating",
        "work_before_break_time", "ergonomic_familiarity", "recurring_chronic_pain", "pain_location",
        "posture_most_used", "posture_switch_to", "posture_change_frequency", "seat_continuously_change_freq",
        *[f"posture_list_familarity_{i}" for i in range(1, 13)],
        *[f"hypermobility_matrix_{i}" for i in range(1, 6)],
    ],
    "anthropometry": ["participant_id"],
    "chair_selection_pre": [
        "participant_id",
        *[f"pre_chair_familiarity_check_{i}" for i in range(1, 12)],
        "pre_selected_chair_code",
        *[f"pre_selection_basis_{i}" for i in range(1, 11)],
        "pre_selection_basis", "pre_selection_confidence_3_3",
    ],
    "chair_selection_post": [
        "participant_id",
        "post_selected_chair_code",
        *[f"post_selection_basis_{i}" for i in range(1, 11)],
        "post_selection_basis", "post_selection_confidence_3_3",
        "selection_changed", "selection_change_reason",
    ],
    "chair_trials": [
        "participant_id", "headrest_dependency", "headrest_dependency_user",
        "interfere", "interfere_location_selected_choice", "interfere_severity",
        "elbow_fit", "seat_depth_fit", "seat_width_fit", "seat_height_fit",
        "backrest_width_fit", "backrest_height_fit", "headrest_height_fit", "made_for_me_fit",
        *[f"geometric_fit_{i}" for i in range(1, 8)],
        "discomfort_seat_3_3", "discomfort_back_3_3", "overall_comfort_3_3",
        "chair_aesthetics_1", "chair_aesthetics_2", "chair_aesthetics_30",
        "expected_vs_experienced_seat_3_3", "expected_vs_experienced_back_3_3",
        "expected_vs_experienced_overall_3_3", "prolonged_willingness_3_3",
        *[f"improve_suggestion_{i}" for i in range(1, 8)],
        *[f"improve_suggestion_loc_{i}" for i in range(1, 7)],
    ],
}
# Join views inherit from their base tables
_CODEBOOK_TABLE_MAP["participant_overview"] = (
    _CODEBOOK_TABLE_MAP["demographics_health"]
    + _CODEBOOK_TABLE_MAP["chair_selection_pre"]
    + _CODEBOOK_TABLE_MAP["chair_selection_post"]
)
_CODEBOOK_TABLE_MAP["chair_trials_analysis"] = _CODEBOOK_TABLE_MAP["chair_trials"]

JOIN_VIEWS = {
    "chair_trials_analysis": """
        SELECT
            ct.*,
            cl.chair_id, cl.chair_model, cl.chair_variant, cl.chair_name_display
        FROM chair_trials ct
        LEFT JOIN chair_lookup cl ON ct.chair_code = cl.chair_code
    """,
    "participant_overview": """
        SELECT
            p.participant_id,
            d.sex_assigned_at_birth,
            d.ethnicity,
            d.age_years,
            d.daily_physical_activity,
            d.recurring_chronic_pain,
            d.general_health_rating,
            d.ergonomic_familiarity,
            d.weekly_desk_hours_time,
            d.work_before_break_time,
            d.seat_continuously_change_freq,
            d.posture_most_used,
            d.posture_change_frequency,
            a.body_mass_kg,
            a.bmi_kg_m2,
            a.body_fat_percent,
            a.skeletal_muscle_percent,
            a.hip_breadth_sitting_cm,
            a.popliteal_height_sitting_cm,
            a.buttock_popliteal_length_cm,
            a.elbow_height_sitting_cm,
            a.sitting_height_erect_cm,
            pre.pre_selected_chair_code,
            pre.pre_selection_basis,
            pre.pre_selection_confidence_3_3,
            pre.pre_selection_basis_1,
            pre.pre_selection_basis_2,
            pre.pre_selection_basis_3,
            pre.pre_selection_basis_4,
            pre.pre_selection_basis_5,
            pre.pre_chair_familiarity_check_1,
            pre.pre_chair_familiarity_check_2,
            post.post_selected_chair_code,
            post.post_selection_basis,
            post.post_selection_confidence_3_3,
            post.post_selection_basis_1,
            post.post_selection_basis_2,
            post.post_selection_basis_3,
            post.post_selection_basis_4,
            post.post_selection_basis_5,
            post.selection_changed,
            post.selection_change_reason
        FROM participants p
        LEFT JOIN demographics_health d ON p.participant_id = d.participant_id
        LEFT JOIN anthropometry a ON p.participant_id = a.participant_id
        LEFT JOIN chair_selection_pre pre ON p.participant_id = pre.participant_id
        LEFT JOIN chair_selection_post post ON p.participant_id = post.participant_id
    """,
}

@st.cache_resource
def get_engine():
    try:
        engine = create_engine(DB_URL)
        # Test connection
        with engine.connect() as conn:
            conn.execute(text("SELECT 1"))
        return engine
    except Exception as e:
        st.error(f"❌ Database connection failed: {e}\n\nPlease ensure the ETL pipeline has been run and Neon is populated.")
        return None

@st.cache_data(show_spinner=False)
def get_columns(table_name: str) -> List[str]:
    engine = get_engine()
    if engine is None:
        # Return fallback columns from TABLE_CONFIG
        return TABLE_CONFIG.get(table_name, {}).get("default_columns", ["error: no data"])
    
    try:
        if table_name in JOIN_VIEWS:
            sql = text(f"SELECT * FROM ({JOIN_VIEWS[table_name]}) q LIMIT 0")
        else:
            sql = text(f'SELECT * FROM "{table_name}" LIMIT 0')
        with engine.connect() as conn:
            df = pd.read_sql(sql, conn)
        return list(df.columns)
    except Exception:
        # Fallback to configured columns
        return TABLE_CONFIG.get(table_name, {}).get("default_columns", [])

@st.cache_data(show_spinner=False)
def get_distinct_values(table_name: str, column_name: str, limit: int = 200):
    engine = get_engine()
    if engine is None:
        return []
    
    try:
        if table_name in JOIN_VIEWS:
            sql = text(f'''
                SELECT DISTINCT "{column_name}"
                FROM ({JOIN_VIEWS[table_name]}) q
                WHERE "{column_name}" IS NOT NULL
                ORDER BY 1
                LIMIT {limit}
            ''')
        else:
            sql = text(f'''
                SELECT DISTINCT "{column_name}"
                FROM "{table_name}"
                WHERE "{column_name}" IS NOT NULL
                ORDER BY 1
                LIMIT {limit}
            ''')
        with engine.connect() as conn:
            result = conn.execute(sql).fetchall()
        return [r[0] for r in result]
    except Exception:
        return []


def build_query(table_name: str, selected_columns: List[str], filters: dict, row_limit: int):
    where_clauses = []
    params = {}

    for idx, (col, meta) in enumerate(filters.items()):
        if meta["mode"] == "multiselect" and meta["value"]:
            placeholders = []
            for j, value in enumerate(meta["value"]):
                key = f"p_{idx}_{j}"
                placeholders.append(f":{key}")
                params[key] = value
            where_clauses.append(f'"{col}" IN ({", ".join(placeholders)})')
        elif meta["mode"] == "contains" and meta["value"]:
            key = f"p_{idx}"
            where_clauses.append(f'CAST("{col}" AS TEXT) ILIKE :{key}')
            params[key] = f'%{meta["value"]}%'
        elif meta["mode"] == "minmax":
            min_val = meta.get("min")
            max_val = meta.get("max")
            if min_val is not None:
                key = f"pmin_{idx}"
                where_clauses.append(f'"{col}" >= :{key}')
                params[key] = min_val
            if max_val is not None:
                key = f"pmax_{idx}"
                where_clauses.append(f'"{col}" <= :{key}')
                params[key] = max_val

    select_clause = ", ".join([f'"{c}"' for c in selected_columns]) if selected_columns else "*"
    where_sql = f"WHERE {' AND '.join(where_clauses)}" if where_clauses else ""

    if table_name in JOIN_VIEWS:
        sql = text(f'''
            SELECT {select_clause}
            FROM ({JOIN_VIEWS[table_name]}) q
            {where_sql}
            LIMIT {row_limit}
        ''')
    else:
        sql = text(f'''
            SELECT {select_clause}
            FROM "{table_name}"
            {where_sql}
            LIMIT {row_limit}
        ''')

    return sql, params


def dataframe_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="query_results", index=False)
    return output.getvalue()


st.title("FitCheck Query Portal")
st.caption("Filter FitCheck study data, preview results, and download CSV or Excel files.")

with st.sidebar:
    # st.header("Connection")
    # st.code(DB_URL.replace(DB_URL.split("@")[-1], "***@***") if "@" in DB_URL else DB_URL)
    st.header("Query Setup")
    data_source = st.selectbox(
        "Choose data source",
        options=[
            "participants",
            "demographics_health",
            "anthropometry",
            "chair_selection_pre",
            "chair_selection_post",
            "chair_lookup",
            "chair_feature_bounds_wide",
            "chair_trials",
            "participant_overview",
            "chair_trials_analysis",
        ],
    )
    st.markdown(
        "**Column scale notes**\n"
        "- `_3_3` columns: −3 to +3 (e.g. comfort, confidence, expected vs experienced)\n"
        "- `geometric_fit_*` columns: −3 to +3 (too small → too large)\n"
        "- `*_familiarity*` / `*_basis_*` columns: binary 0/1 (unchecked / checked)\n"
        "- Measurement columns: cm, kg, %, or minutes as indicated in the name"
    )

    all_columns = get_columns(data_source)
    default_columns = TABLE_CONFIG.get(data_source, {}).get("default_columns", all_columns[: min(8, len(all_columns))])
    selected_columns = st.multiselect("Columns to include", all_columns, default=default_columns)
    row_limit = st.number_input("Max rows to return", min_value=10, max_value=100000, value=990, step=10)

    st.divider()
    st.header("Cleaned Export")
    if CLEANED_EXPORT_PATH.exists():
        _export_mtime = CLEANED_EXPORT_PATH.stat().st_mtime
        import datetime
        _export_ts = datetime.datetime.fromtimestamp(_export_mtime).strftime("%Y-%m-%d %H:%M")
        st.caption(f"Last generated: {_export_ts}")
        with open(CLEANED_EXPORT_PATH, "rb") as _f:
            st.download_button(
                label="Download fitcheck_cleaned_export.xlsx",
                data=_f.read(),
                file_name="fitcheck_cleaned_export.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    else:
        st.warning(
            "fitcheck_cleaned_export.xlsx not found.\n\n"
            f"Expected at: `{CLEANED_EXPORT_PATH}`\n\n"
            "Run the ETL pipeline to generate it."
        )

    st.divider()
    with st.expander("📖 Variable Codebook", expanded=False):
        _cb_search = st.text_input(
            "Search variable name or question",
            placeholder="e.g. sex, comfort, fit...",
            key="codebook_search",
        )
        _cb_vars = _CODEBOOK_TABLE_MAP.get(data_source, list(CODEBOOK.keys()))
        _cb_vars_in_table = [v for v in _cb_vars if v in CODEBOOK]

        if _cb_search.strip():
            _q = _cb_search.strip().lower()
            _cb_vars_in_table = [
                v for v in _cb_vars_in_table
                if _q in v.lower() or _q in CODEBOOK[v]["question"].lower()
            ]

        if not _cb_vars_in_table:
            st.caption("No matching variables found." if _cb_search.strip() else "No coded variables for this data source.")
        else:
            st.caption(f"{len(_cb_vars_in_table)} variable(s) shown")
            for var in _cb_vars_in_table:
                entry = CODEBOOK[var]
                st.markdown(f"**`{var}`** — *{entry['question']}*")
                if entry["options"]:
                    rows = [{"Code": k, "Label": v} for k, v in entry["options"].items()]
                    st.dataframe(
                        pd.DataFrame(rows),
                        hide_index=True,
                        use_container_width=True,
                        height=min(35 * len(rows) + 38, 250),
                    )
                else:
                    st.caption(f"  Type: {entry['type']} (no coded values)")
                st.markdown("---")

st.subheader("Filters")
filterable_columns = TABLE_CONFIG.get(data_source, {}).get("filters", all_columns[: min(6, len(all_columns))])
filters = {}

cols = st.columns(2)
for i, column_name in enumerate(filterable_columns):
    with cols[i % 2]:
        st.markdown(f"**{column_name}**")
        mode = st.selectbox(
            f"Filter type for {column_name}",
            ["none", "multiselect", "contains", "minmax"],
            key=f"mode_{column_name}",
        )
        if mode == "multiselect":
            options = get_distinct_values(data_source, column_name)
            value = st.multiselect(f"Values for {column_name}", options, key=f"vals_{column_name}")
            filters[column_name] = {"mode": mode, "value": value}
        elif mode == "contains":
            value = st.text_input(f"Contains text for {column_name}", key=f"txt_{column_name}")
            filters[column_name] = {"mode": mode, "value": value}
        elif mode == "minmax":
            min_val = st.text_input(f"Min for {column_name}", key=f"min_{column_name}")
            max_val = st.text_input(f"Max for {column_name}", key=f"max_{column_name}")
            filters[column_name] = {
                "mode": mode,
                "min": None if min_val == "" else min_val,
                "max": None if max_val == "" else max_val,
            }

run_query = st.button("Run Query", type="primary")

if run_query:
    engine = get_engine()
    if engine is None:
        st.error("❌ Cannot run query: Database not connected. Please wait for the ETL pipeline to complete and refresh the page.")
    else:
        try:
            sql, params = build_query(data_source, selected_columns, filters, row_limit)
            with engine.connect() as conn:
                result_df = pd.read_sql(sql, conn, params=params)

            st.success(f"Query returned {len(result_df)} rows.")
            st.dataframe(result_df, use_container_width=True, height=500)

            c1, c2, c3 = st.columns(3)
            with c1:
                st.download_button(
                    label="Download CSV",
                    data=result_df.to_csv(index=False).encode("utf-8"),
                    file_name=f"{data_source}_query_results.csv",
                    mime="text/csv",
                )
            with c2:
                st.download_button(
                    label="Download Excel",
                    data=dataframe_to_excel_bytes(result_df),
                    file_name=f"{data_source}_query_results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            with c3:
                full_sql_preview = str(sql)
                st.download_button(
                    label="Download SQL",
                    data=full_sql_preview.encode("utf-8"),
                    file_name=f"{data_source}_query.sql",
                    mime="text/plain",
                )

        except Exception as e:
            st.error(f"Query failed: {e}")

st.markdown("---")
# st.subheader("Suggested next improvements")
# st.write(
#     "Add login/authentication, save common queries, include charts for summary views, and offer cross-table participant filtering."
# )

# st.code(
#     """
# # Example run command
# streamlit run fitcheck_query_app.py

# # Example environment variable (Windows PowerShell)
# $env:POSTGRES_URL = 'postgresql+psycopg2://username:password@localhost:5432/fitcheck'
#     """.strip(),
#     language="bash",
# )
