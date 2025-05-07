import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import plotly.express as px
from datetime import timedelta
from openai import OpenAI
from helper_exports import create_powerpoint, create_word_doc

st.set_page_config(page_title="Scenario Simulation Assistant", layout="wide")
st.title("Primavera Scenario Simulation Assistant")

# Upload files
st.header("Step 1: Upload Your Primavera Schedules")
col1, col2 = st.columns(2)
with col1:
    before_file = st.file_uploader("Upload BASELINE schedule (Excel)", type=["xlsx"], key="before")
with col2:
    after_file = st.file_uploader("Upload UPDATED schedule (Excel)", type=["xlsx"], key="after")

# Process and compare
if before_file and after_file:
    df_before = pd.read_excel(before_file)
    df_after = pd.read_excel(after_file)

    merge_key = "Activity ID"
    df_merged = pd.merge(df_before, df_after, on=merge_key, suffixes=('_before', '_after'))
    compare_cols = ["Start", "Finish", "Total Float"]
    display_cols = [merge_key, "Activity Name_before"] + [f"{col}_before" for col in compare_cols] + [f"{col}_after" for col in compare_cols]

    changed_rows = df_merged[
        (df_merged["Start_before"] != df_merged["Start_after"]) |
        (df_merged["Finish_before"] != df_merged["Finish_after"]) |
        (df_merged["Total Float_before"] != df_merged["Total Float_after"])
    ]

    st.subheader("Changed Activities")
    st.dataframe(changed_rows[display_cols])

    # Float chart
    st.subheader("Float Change Chart")
    changed_rows["Float Change"] = changed_rows["Total Float_after"] - changed_rows["Total Float_before"]
    fig_float, ax = plt.subplots(figsize=(10, 4))
    ax.bar(changed_rows["Activity Name_before"], changed_rows["Float Change"])
    ax.set_title("Float Change (After - Before)")
    ax.set_ylabel("Days")
    ax.tick_params(axis='x', rotation=45)
    st.pyplot(fig_float)

    # Milestone movement
    st.subheader("Milestone Movement Plot")
    milestone_mask = changed_rows["Activity Name_before"].str.contains("milestone", case=False)
    milestone_df = changed_rows[milestone_mask][[merge_key, "Activity Name_before", "Finish_before", "Finish_after"]].copy()

    if not milestone_df.empty:
        milestone_df = pd.melt(
            milestone_df,
            id_vars=["Activity Name_before"],
            value_vars=["Finish_before", "Finish_after"],
            var_name="Version",
            value_name="Date"
        )
        fig_milestone = px.line(milestone_df, x="Date", y="Activity Name_before", color="Version", markers=True)
        st.plotly_chart(fig_milestone)
    else:
        st.info("No milestone changes detected.")
        fig_milestone, ax = plt.subplots(figsize=(6, 3))
        ax.text(0.5, 0.5, "No milestone changes detected", fontsize=12, ha='center', va='center')
        ax.set_axis_off()

    # GPT Summary
    st.subheader("AI Impact Summary (Structured & Bullet Format)")
    client = OpenAI(api_key=st.secrets["openai"]["api_key"])

    summary_prompt = f"""
You are a senior delay analyst reporting on Primavera schedule changes. Use the exact structure and markdown format below. Be concise and use bullet points for all sections except the Overview.

DATA:
{changed_rows[display_cols].to_string(index=False)}

FORMAT:
### 1. Overview
Summarize what happened in 2–3 lines.

### 2. Impacted Activities
- List key activities and how much they shifted

### 3. Critical Path / Float Changes
- Mention float reductions or new critical paths

### 4. Milestone Impacts
- Identify milestone delays with old/new dates

### 5. Recommendations
- Suggest 2–3 actions to mitigate or recover

Output using the exact format and bullet structure above.
"""

    if st.button("Generate AI Summary"):
        response = client.chat.completions.create(
            model="gpt-4",
            messages=[{"role": "user", "content": summary_prompt}]
        )
        gpt_output = response.choices[0].message.content
        st.markdown(gpt_output)

        # Export buttons
        st.success("Generate reports with explanation and visuals:")
        col1, col2 = st.columns(2)

        with col1:
            pptx_path = create_powerpoint(gpt_output, fig_float, fig_milestone)
            with open(pptx_path, "rb") as f:
                st.download_button("Download PowerPoint", f, file_name="Delay_Impact_Analysis.pptx")

        with col2:
            word_path = create_word_doc(gpt_output, fig_float, fig_milestone)
            with open(word_path, "rb") as f:
                st.download_button("Download Word Report", f, file_name="Delay_Impact_Report.docx")
