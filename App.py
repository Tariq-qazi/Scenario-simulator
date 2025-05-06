import streamlit as st
import pandas as pd
from datetime import timedelta
import matplotlib.pyplot as plt
import plotly.express as px
import openai

# App title
st.set_page_config(page_title="Scenario Simulation Assistant", layout="wide")
st.title("Primavera Scenario Simulation Assistant")

# Upload the Excel files
st.header("Step 1: Upload Your Primavera Schedules")
col1, col2 = st.columns(2)
with col1:
    before_file = st.file_uploader("Upload BASELINE schedule (Excel)", type=["xlsx"], key="before")
with col2:
    after_file = st.file_uploader("Upload UPDATED schedule (Excel)", type=["xlsx"], key="after")

# Load and compare
if before_file and after_file:
    df_before = pd.read_excel(before_file)
    df_after = pd.read_excel(after_file)

    st.success("Both files loaded. Proceeding with analysis...")

    merge_key = "Activity ID"
    if merge_key not in df_before.columns:
        merge_key = st.selectbox("Select ID column for merging", df_before.columns)

    # Merge and compare
    df_merged = pd.merge(df_before, df_after, on=merge_key, suffixes=('_before', '_after'))
    compare_cols = ["Start", "Finish", "Total Float"]

    changed_rows = df_merged[
        (df_merged["Start_before"] != df_merged["Start_after"]) |
        (df_merged["Finish_before"] != df_merged["Finish_after"]) |
        (df_merged["Total Float_before"] != df_merged["Total Float_after"])
    ]

    st.subheader("Changed Activities")
    display_cols = [merge_key, "Activity Name_before"] + [f"{col}_before" for col in compare_cols] + [f"{col}_after" for col in compare_cols]
    st.dataframe(changed_rows[display_cols])

    # Float change chart
    st.subheader("Float Change Chart")
    changed_rows["Float Change"] = changed_rows["Total Float_after"] - changed_rows["Total Float_before"]
    fig, ax = plt.subplots(figsize=(10, 4))
    ax.bar(changed_rows["Activity Name_before"], changed_rows["Float Change"])
    ax.set_title("Float Change (After - Before)")
    ax.set_ylabel("Days")
    ax.tick_params(axis='x', rotation=45)
    st.pyplot(fig)

    # Milestone movement chart
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
        fig2 = px.line(milestone_df, x="Date", y="Activity Name_before", color="Version", markers=True)
        st.plotly_chart(fig2)
    else:
        st.info("No milestone changes detected.")

    # GPT summary
    st.subheader("AI Impact Summary")
    openai.api_key = st.secrets["openai"]["api_key"]

    summary_prompt = f"""
    Analyze the following schedule changes:
    {changed_rows[display_cols].to_string(index=False)}

    Provide a 2-paragraph summary of:
    - Key delays and affected activities
    - Impacts on float and milestones
    - Any possible critical path implications
    """

    if st.button("Generate Summary"):
        response = openai.ChatCompletion.create(
            model="gpt-4",
            messages=[{"role": "user", "content": summary_prompt}]
        )
        st.markdown(response["choices"][0]["message"]["content"])
