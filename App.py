import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import plotly.express as px
from datetime import timedelta
from openai import OpenAI
from helper_exports import create_powerpoint, create_word_doc  # External helper file

st.set_page_config(page_title="Scenario Simulation Assistant", layout="wide")
st.title("Primavera Scenario Simulation Assistant")

# Step 1: Upload Files
st.header("Step 1: Upload Your Primavera Schedules")
col1, col2 = st.columns(2)
with col1:
    before_file = st.file_uploader("Upload BASELINE schedule (Excel)", type=["xlsx"], key="before")
with col2:
    after_file = st.file_uploader("Upload UPDATED schedule (Excel)", type=["xlsx"], key="after")

# Step 2: Process and Compare
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

    # Step 3: Float Change Chart
    st.subheader("Float Change Chart")
    changed_rows["Float Change"] = changed_rows["Total Float_after"] - changed_rows["Total Float_before"]
    fig_float, ax = plt.subplots(figsize=(10, 4))
    ax.bar(changed_rows["Activity Name_before"], changed_rows["Float Change"])
    ax.set_title("Float Change (After - Before)")
    ax.set_ylabel("Days")
    ax.tick_params(axis='x', rotation=45)
    st.pyplot(fig_float)

    # Step 4: Milestone Movement Chart
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
        fig_milestone = plt.figure()  # empty placeholder for export

    # Step 5: GPT Structured Summary
    st.subheader("AI Impact Summary (Structured)")
    client = OpenAI(api_key=st.secrets["openai"]["api_key"])

    summary_prompt = f"""
You are a scheduling analyst. Summarize the following Primavera schedule changes using short, clear sections with bullet points.

DATA:
{changed_rows[display_cols].to_string(index=False)}

FORMAT:
**1. Overview**
- What changed and why

**2. Impacted Activities**
- List delayed or shifted activities and by how much

**3. Critical Path / Float Changes**
- Mention float changes and any new critical path exposure

**4. Milestone Impacts**
- Identify delayed milestones and how far they shifted

**5. Recommendations**
- Suggest possible actions or mitigations (add crews, resequence, etc.)

Respond in clean markdown format.
"""

    if st.button("Generate AI Summary"):
        response = client.chat.completions.create(
            model="gpt-4",
            messages=[{"role": "user", "content": summary_prompt}]
        )
        gpt_output = response.choices[0].message.content
        st.markdown(gpt_output)

        # Export Options
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
