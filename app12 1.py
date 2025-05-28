import streamlit as st
import pandas as pd
import re
import random
from streamlit.components.v1 import html
from io import BytesIO
import openpyxl

st.set_page_config(page_title="Skill Matcher", layout="wide")

# --- Init session state ---
if "uploaded" not in st.session_state:
    st.session_state.uploaded = False
if "bench_demand_file" not in st.session_state:
    st.session_state.bench_demand_file = None
if "subcon_file" not in st.session_state:
    st.session_state.subcon_file = None
if "master_skill_file" not in st.session_state:
    st.session_state.subcon_file = None


def validate_file(uploaded_file, file_description):
    allowed_types = [
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","xlsx"
]
    if uploaded_file is None:
        return False
    if uploaded_file.type not in allowed_types:
        st.error(
            f"❌ The file '{uploaded_file.name}' is not supported. "
            f"Please upload a valid Excel file (.xlsx) for {file_description}."
        )
        return False
    return True

# --- Upload Page ---
def upload_page():
    st.title("📤 Upload Required Excel Files")
    st.markdown("Please upload the required files to continue:")

    bench_demand = st.file_uploader("Upload Bench & Demand Excel File", key="bench_demand")
    is_bench_demand_valid = validate_file(bench_demand, "Bench & Demand")
    subcon = st.file_uploader("Upload Sub-Con Candidate Report Excel File", key="Subcon")
    is_subcon_valid = validate_file(subcon, "Sub-Con Candidate Report")
    master_skill = st.file_uploader("Master Skills Excel File", key="master_skill")
    is_master_skill_valid = validate_file(master_skill, "Master Skills")



    if is_bench_demand_valid and is_subcon_valid and is_master_skill_valid:
        st.session_state.bench_demand_file = bench_demand
        st.session_state.subcon_file = subcon
        st.session_state.master_skill_file = master_skill
        st.session_state.uploaded = True
        st.rerun()

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine = openpyxl) as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data

# --- Main App Page ---
def main_app():
    st.title("🧠 RMG Skill Matcher")
    st.button("🔄 Re-upload Files", on_click=lambda: reset_files())

    try:
        xls = pd.ExcelFile(st.session_state.bench_demand_file)
        bench_df = xls.parse('Bench Base')
        demand_df = xls.parse('Demand Base')
        # master_skill_df = xls.parse('master skill1')
        xls0 = pd.ExcelFile(st.session_state.master_skill_file)
        master_skill_df = xls0.parse('MasterList')
        xls1 = pd.ExcelFile(st.session_state.subcon_file)
        subcon_df = xls1.parse('Engineering')

        st.header("🔍 Filter Criteria")

        col1, col2, col3, col4 = st.columns([2, 2, 2, 2])
        with col1:
            practice_options = sorted(bench_df["Practice"].dropna().unique().tolist())
            selected_practice = st.multiselect("Practice", practice_options)
        bench_filtered = bench_df.copy()
        if selected_practice:
            bench_filtered = bench_filtered[bench_filtered["Practice"].isin(selected_practice)]

        with col2:
            sub_practice_options = sorted(bench_filtered["Sub Practice"].dropna().unique().tolist())
            selected_sub_practice = st.multiselect("Sub Practice", sub_practice_options)
        if selected_sub_practice:
            bench_filtered = bench_filtered[bench_filtered["Sub Practice"].isin(selected_sub_practice)]

        with col3:
            grade_options = sorted(bench_filtered["Grade"].dropna().unique().tolist())
            selected_grade = st.multiselect("Grade", grade_options)
        if selected_grade:
            bench_filtered = bench_filtered[bench_filtered["Grade"].isin(selected_grade)]

        with col4: 
            skill_grouping_options = sorted(bench_filtered["Skill Grouping"].dropna().unique().tolist())
            selected_skill_grouping = st.multiselect("Skill Grouping", skill_grouping_options)
        if selected_skill_grouping:
            bench_filtered = bench_filtered[bench_filtered["Skill Grouping"].isin(selected_skill_grouping)]

        row2_col1, row2_col2 = st.columns([3, 3])
        with row2_col1:
            search_name = st.text_input("Search by Employee Name")
        if search_name:
            bench_filtered = bench_filtered[
                bench_filtered["EmployeeName"].str.lower().str.startswith(search_name.lower())
            ]

        with row2_col2:
            search_skill = st.text_input("Search by Skill")
        if search_skill:
            bench_filtered = bench_filtered[
                bench_filtered["Skill"].str.lower().str.contains(search_skill.lower(), na=False)
            ]

        st.subheader("📋 Filtered Bench Base")
        display_df = bench_filtered[
            ["Region", "Country","Practice", "Sub Practice", "LDAP ID", "EmployeeName", "Email", "Grade", "Skill Grouping", "Skill"]
        ].copy()
        # display_df["ID"] = display_df["ID"].astype(str)
        display_df["LDAP ID"] = display_df["LDAP ID"].astype(str)
        st.dataframe(display_df, use_container_width=True)

        # Extract Skills from 'Skill' column
        st.subheader("🧩 Extracted Skills based on Filters")
        skill_set = set()
        for skill_text in bench_filtered["Skill"].dropna():
            split_skills = re.split(r'[,:;]', skill_text)
            skill_set.update([s.strip().lower() for s in split_skills if s.strip()])
        skill_list = sorted(skill_set)

        if skill_list:
            html_string = f"""
            <div style="
                max-height: 200px;
                overflow-y: auto;
                border: 1px solid #ccc;
                padding: 10px;
                border-radius: 10px;
                background-color: #f9f9f9;">
            """
            for skill in skill_list:
                html_string += f"""
                <span style='
                    display: inline-block;
                    background-color: #dbeafe;
                    color: #1e3a8a;
                    font-size: 14px;
                    padding: 5px 10px;
                    margin: 5px;
                    border-radius: 20px;
                    border: 1px solid #93c5fd;
                '>{skill}</span>
                """
            html_string += "</div>"
            html(html_string, height=250)
        else:
            st.info("No skills available to display.")
        selected_options = st.slider(
            "Select the Filters Range",
            min_value=0,
            max_value=100,
            value=(0, 100),
            step=10
            )
        # Buttons
        btn_col1, btn_col2, btn_col3, btn_col4 = st.columns([1, 1, 1, 1])
        with btn_col2:
            search_demand_clicked = st.button("🔍 Search Matching Demand")
        with btn_col3:
            search_subcon_clicked = st.button("🔎 Search Matching Sub-Con")

        # Matching logic
        all_skills = set()
        for skills in bench_filtered["Skill"].dropna():
            for skill in re.split(r'[,:;]', skills):
                skill = skill.strip().lower()
                if skill:
                    all_skills.add(skill)

        def compute_match_percent(skills_text):
            if pd.isna(skills_text) or not all_skills:
                return 0
            match_count = sum(1 for skill in all_skills if skill in skills_text.lower())
            return round((match_count / len(all_skills)) * 100, 2) if all_skills else 0

        # --- Matching logic for Demand search ---
        
        if search_demand_clicked:
            st.subheader("🎯 Matched Demand Base")
            
            bench_info_cols = ["LDAP ID", "EmployeeName", "Email", "Skill"]
            bench_info = bench_filtered[bench_info_cols].reset_index(drop=True)

            # Two alternating row colors
            colors = ['#f0f8ff', '#e6f7ff']
            
            all_matched_demand_list = []

            # Loop through each bench employee
            for idx, bench_row in bench_info.iterrows():
                ldap_id = bench_row['LDAP ID']
                employee_name = bench_row['EmployeeName']
                email = bench_row['Email']
                skills_raw = bench_row["Skill"]

                if pd.isna(skills_raw):
                    st.warning(f"⚠️ No skills provided for {employee_name}.")
                    continue
                emp_skills_list = (s.strip().lower() for s in re.split(r'[,:;]', skills_raw) if s.strip())

                def get_skills_list(emp_skills_list):
                    if not emp_skills_list:
                        return []

                    # Remove NaN and convert all to lowercase stripped strings
                    emp_skills_list = [str(skill).strip().lower() for skill in emp_skills_list if pd.notna(skill)]

                    matched_skills = set(emp_skills_list)

                    for _, row in master_skill_df.iterrows():
                        # Skip rows where 'Skills' is NaN
                        if pd.isna(row['Skills']):
                            continue

                        primary_skill = str(row['Skills']).strip().lower()

                        # Handle NaN in 'Alias' gracefully
                        alias_str = str(row['Alias']) if pd.notna(row['Alias']) else ""
                        alias = [s.strip().lower() for s in re.split(r'[,:;]', alias_str) if s.strip()]

                        all_master_skills = [primary_skill] + alias

                        # Match: if any employee skill is in the master skills
                        if any(emp_skill in all_master_skills for emp_skill in emp_skills_list):
                            matched_skills.add(primary_skill)
                            matched_skills.update(alias)
                    # Return sorted, deduplicated list
                    return sorted(matched_skills)
                
                emp_skills_list = get_skills_list(emp_skills_list)
                
                def compute_match_percent_and_skills(mandatory_text):
                    if pd.isna(mandatory_text) or not emp_skills_list:
                        return 0, [], emp_skills_list

                    # Normalize and split the mandatory skills
                    mandatory_list = [s.strip().lower() for s in str(mandatory_text).split(",") if s.strip()]
                    
                    # Normalize employee skills
                    normalized_emp_skills = (s.strip().lower() for s in emp_skills_list if pd.notna(s))

                    matched = [s for s in mandatory_list if s in normalized_emp_skills]
                        
                    match_percent = round((len(matched) / len(mandatory_list)) * 100, 2) if mandatory_list else 0
                    return match_percent, matched

                matched_demand = demand_df.copy()
                matched_demand[["Match %", "Matching Skills"]] = matched_demand["Mandatory Skills"].apply(
                    lambda x: compute_match_percent_and_skills(x)
                ).apply(pd.Series)

                # Filter out rows with no match
                matched_demand = matched_demand[matched_demand["Matching Skills"].notnull()]
                matched_demand = matched_demand[matched_demand["Match %"] > 0]
                matched_demand = matched_demand.sort_values(by="Match %", ascending=False)

                # After computing matched_demand for this employee:
                if not matched_demand.empty:
                    # Add bench employee details to each demand row
                    matched_demand["LDAP ID"] = ldap_id
                    matched_demand["EmployeeName"] = employee_name
                    matched_demand["Email"] = email
                    matched_demand["Skill"] = skills_raw

                    # Append this employee's matched_demand to the list
                    all_matched_demand_list.append(matched_demand)

                else:
                    pass
            combined_display = []
            # After the loop finishes, combine all employee matching demand data
            if all_matched_demand_list:
                combined_df = pd.concat(all_matched_demand_list, ignore_index=True)

                # Select columns to display (optional: reorder/filter as you like)
                demand_cols_to_show = ["LDAP ID", "EmployeeName", "Email", "Skill", "ID", "Client", "Project Name", 
                                        "Mandatory Skills", "Matching Skills", "Match %"]
                    
                combined_display = combined_df[demand_cols_to_show].sort_values(by="Match %", ascending=False)

                st.markdown("## Combined Matching Demand for All Employees")

                    # Display combined results

            filtered_combined_display = combined_display[combined_display["Match %"].between(*selected_options)]

            # Show warning if no data matches
            if filtered_combined_display.empty:
                st.warning("⚠️ No matches found in the selected range.")
            else:
                # Convert to Excel
                excel_data = to_excel(filtered_combined_display)

                # Download button (only shown when data exists)
                st.download_button(
                    label="📥 Download Combined Demand Match Excel File",
                    data=excel_data,
                    file_name='combined_bench_demand_match.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )

            # Display DataFrame (even if empty, optionally)
            st.dataframe(filtered_combined_display, use_container_width=True, hide_index=True)





        # --- Matching logic for Sub-con search ---
        if search_subcon_clicked:
            st.subheader("🤝 Matched Sub-Con Candidates")

            if "Skill" not in subcon_df.columns:
                st.error("❌ 'Skill' column not found in Sub-Con Excel file.")
            else:
                bench_info_cols = ["LDAP ID", "EmployeeName", "Email", "Skill"]
                bench_info = bench_filtered[bench_info_cols].reset_index(drop=True)

                # Same alternating colors as Demand
                colors = ['#f0f8ff', '#e6f7ff']
                all_matched_subcon_list = []
                for idx, bench_row in bench_info.iterrows():
                    ldap_id = bench_row['LDAP ID']
                    employee_name = bench_row['EmployeeName']
                    email = bench_row['Email']
                    skills_raw = bench_row["Skill"]

                    if pd.isna(skills_raw):
                        st.warning(f"⚠️ No skills provided for {employee_name}.")
                        continue

                    emp_skills_list = [s.strip().lower() for s in re.split(r'[,:;]', skills_raw) if s.strip()]

                    def get_skills_list(emp_skills_list):
                        if not emp_skills_list:
                            return []

                        matched_skills = set(emp_skills_list)

                        for _, row in master_skill_df.iterrows():
                            primary_skill = str(row['Skills']).strip().lower()
                            # all_master_skills = [primary_skill] + secondary_skills
                            alias = [s.strip().lower() for s in re.split(r'[,:;]', str(row['Alias'])) if s.strip()]
                            all_master_skills = [primary_skill]+[alias]

                            # Match: if any employee skill is in the master skills
                            if any(emp_skill in all_master_skills for emp_skill in emp_skills_list):
                                matched_skills.add(primary_skill)
                                matched_skills.update(alias)

                        # Return sorted, deduplicated list
                        return sorted(matched_skills)
                    
                    emp_skills_list = get_skills_list(emp_skills_list)

                    def compute_match_percent_and_skills(subcon_skill_text):
                        if pd.isna(subcon_skill_text) or not emp_skills_list:
                            return 0, []
                        subcon_skills = [s.strip().lower() for s in re.split(r'[,:;]', subcon_skill_text) if s.strip()]
                        matched = [s for s in emp_skills_list if s in subcon_skills]
                        match_percent = round((len(matched) / len(subcon_skills)) * 100, 2)
                        return match_percent, matched
                    


                    matched_subcon = subcon_df.copy()
                    matched_subcon[["Match %", "Matching Skills"]] = matched_subcon["Skill"].apply(
                        lambda x: compute_match_percent_and_skills(x)
                    ).apply(pd.Series)

                    matched_subcon = matched_subcon[matched_subcon["Match %"] > 0]
                    matched_subcon = matched_subcon.sort_values("Match %", ascending=False)

                    if not matched_subcon.empty:
                        # st.markdown(f"### {employee_name}'s Matching Sub-Con Candidates")

                        matched_subcon["LDAP ID"] = ldap_id
                        matched_subcon["EmployeeName"] = employee_name
                        matched_subcon["Email"] = email
                        matched_subcon["Bench Skill"] = skills_raw

                        all_matched_subcon_list.append(matched_subcon)
                    else:
                        pass
                combined_display = []
                # After the loop finishes, combine all employee matching demand data
                if all_matched_subcon_list:
                    combined_df = pd.concat(all_matched_subcon_list, ignore_index=True)

                    # Select columns to display (optional: reorder/filter as you like)
                    subcon_cols_to_show = ["LDAP ID", "EmployeeName", "Email", "Bench Skill",
                                      "Emp ID", "Consultant Name", "Project Manager", "Client", 
                                      "Skill", "Matching Skills", "Match %"]

                    combined_display = combined_df[subcon_cols_to_show].sort_values(by="Match %", ascending=False)

                    st.markdown("## Combined Matching Subcon for All Employees")

                        # Display combined results

                filtered_combined_display = combined_display[combined_display["Match %"].between(*selected_options)]

                # Show warning if no data matches
                if filtered_combined_display.empty:
                    st.warning("⚠️ No matches found in the selected range.")
                else:
                    # Convert to Excel
                    excel_data = to_excel(filtered_combined_display)

                    # Download button (only shown when data exists)
                    st.download_button(
                        label="📥 Download Combined Subcon Match Excel File",
                        data=excel_data,
                        file_name='combined_bench_demand_match.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )

                # Display DataFrame (even if empty, optionally)
                st.dataframe(filtered_combined_display, use_container_width=True, hide_index=True)

    except Exception as e:
        st.error(f"🚨 Error occurred while reading files: {e}")

# --- Reset function ---
def reset_files():
    st.session_state.uploaded = False
    st.session_state.bench_demand_file = None
    st.session_state.subcon_file = None
    upload_page()

# --- App Router ---
if not st.session_state.uploaded:
    upload_page()
else:
    main_app()
