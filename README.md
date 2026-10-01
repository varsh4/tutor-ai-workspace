AI Tutor Workspace

An interactive Streamlit tutoring workspace for any-subject tutoring, academic problem solving, statistics, mathematics, quantitative reasoning, Excel/data analysis, file assessment, research, and colorful educational visuals.

Academic accuracy architecture

The app uses a hybrid approach:

Deterministic local verification for supported math/statistics tasks (for example regression, correlation, descriptive statistics, z-scores, selected hypothesis tests, equations, percentages, and combinatorics).

Gemini Python code execution for complex or unfamiliar numerical/data-analysis tasks so the model can calculate and verify rather than rely on mental arithmetic.

Gemini Google Search grounding for current/research questions when needed.

Original file inputs are sent to Gemini when practical, while extracted text is also kept as context.

Programmatic graphs are used for numerical/statistical plots; image generation is used for conceptual/colorful teaching visuals.

This design is intended to reduce numerical hallucinations. No AI system can guarantee correctness for every possible task, so graded work should still be reviewed against the course instructions/rubric.

Files

app.py — Streamlit application

requirements.txt — dependencies

.streamlit/config.toml — Streamlit upload/message settings

Streamlit Secrets

Add:

GEMINI_API_KEY = "YOUR_API_KEY"

Optional model settings:

GEMINI_MODEL = "gemini-3.8-flash"
GEMINI_FALLBACK_MODEL = "gemini-3.7-flash"
GEMINI_IMAGE_MODEL = "gemini-3.1-flash-image"
