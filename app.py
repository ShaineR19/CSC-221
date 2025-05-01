# -*- coding: utf-8 -*-
import traceback
import streamlit as st
import pandas as pd
import io
import matplotlib.pyplot as plt
import web_functions as wf
import options4 as opfour

@st.cache_data
def load_data():
    dean_df = wf.readfile()
    unique_df = pd.read_excel('unique_deansDailyCsar_FTE.xlsx')
    fte_tier = pd.read_excel('FTE_Tier.xlsx')
    dean_df.columns = dean_df.columns.str.strip()
    unique_df.columns = unique_df.columns.str.strip()
    return dean_df, unique_df, fte_tier

dean_df, unique_df, fte_tier = load_data()

def save_report(df_full, filename):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_full.to_excel(writer, sheet_name='Full Report', index=False)
    st.download_button("Save Report", output.getvalue(), file_name=filename)


st.title('FTE Report Generator')

menu = [
    "Sec Division Report",
    "Course Enrollment Percentage",
    "FTE by Division",
    "FTE per Instructor",
    "FTE per Course"
]

choice = st.sidebar.radio("Choose Report Option", menu)

if choice == "Sec Division Report":
    st.header("Sec Division Report")
    if 'Sec Divisions' in dean_df.columns:
        division = st.selectbox("Select Division", dean_df['Sec Divisions'].dropna().unique())
        run = st.button("Run Report")
        if run:
            filtered = dean_df[dean_df['Sec Divisions'] == division]
            st.dataframe(filtered.head(10))
            save_report(filtered, f"{division}_Division_Report.xlsx")
    else:
        st.warning("This feature will run when 'Sec Divisions' is available in the dataset.")

elif choice == "Course Enrollment Percentage":
    st.header("Course Enrollment Percentage")
    if 'Sec Name' in dean_df.columns:
        course = st.selectbox("Select Course", dean_df['Sec Name'].dropna().unique())
        run = st.button("Run Report")
        if run:
            filtered = dean_df[dean_df['Sec Name'] == course]
            filtered = filtered.drop_duplicates(subset="Sec Name")
            filtered["Enrollment Percentage"] = filtered.apply(wf.calc_enrollment, axis=1)
            st.dataframe(filtered.head(10))
            filtered['Enrollment Percentage'] = filtered['Enrollment Percentage'].replace('%', '', regex=True).astype(float)
            st.bar_chart(filtered.set_index('Sec Name')['Enrollment Percentage'])
            save_report(filtered, f"{course}_Course_Report.xlsx")
    else:
        st.warning("This feature will run when 'Sec Name' is available in the dataset.")

elif choice == "FTE by Division":
    st.header("FTE by Division")

    if 'Sec Divisions' in dean_df.columns:
        division_input = st.selectbox("Select Division", dean_df['Sec Divisions'].dropna().unique())
        run = st.button("Run Report")

        if run:
            raw_df, orig_total, gen_total = wf.fte_by_div_raw(dean_df, fte_tier, division_input)
            formatted_df = wf.format_fte_output(raw_df, orig_total, gen_total)

            # Add numeric column for plotting
            #formatted_df['Total FTE (Numeric)'] = (
                #formatted_df['Total FTE'].replace('[\$,]', '', regex=True).replace('', '0').astype(float)
            #)

            st.dataframe(formatted_df)

            #plot_df = formatted_df.dropna(subset=['Sec Name', 'Total FTE (Numeric)'])
            #plot_df = plot_df[plot_df['Course Code'] != 'Total'].set_index('Sec Name')

            #if not plot_df.empty:
                #st.bar_chart(plot_df['Total FTE (Numeric)'])

            save_report(formatted_df, f"{division_input}_FTE_Report.xlsx")
    else:
        st.info("Division data not available.")

elif choice == "FTE per Instructor":
    st.header("FTE per Instructor")
    if 'Sec All Faculty Last Names' in dean_df.columns:
        faculty_list = sorted(dean_df['Sec Faculty Info'].dropna().unique())
        instructor = st.selectbox("Select Instructor", faculty_list)

        run = st.button("Run Report")
        if run:
            report_df, orig_fte, gen_fte = wf.generate_faculty_fte_report(dean_df, fte_tier, instructor)
            report_df = report_df.fillna("")

            st.dataframe(report_df)

            filename = opfour.clean_instructor_name(instructor)
            save_report(report_df, filename)

            st.info(f"Total FTE: {orig_fte:.2f}")
            st.info(f"Generated FTE: ${gen_fte:,.2f}")

        else:
            st.warning("Select an Instructor.")
    else:
        st.warning("Instructor name column missing.")

elif choice == "FTE per Course":
    st.header("FTE per Course")
    if 'Sec Name' in unique_df.columns:
        course_name = st.text_input("Enter Course Name (Sec Name)")
        run = st.button("Run Report")
        if run:
            df_result, original_fte, generated_fte = wf.calculate_fte_by_course(dean_df, fte_tier, course_name)
            if df_result is not None:
                st.dataframe(df_result)
                save_report(df_result, f"{course_name}_FTE_Report.xlsx")
                st.success(f"Report generated for {course_name}")
                st.bar_chart(df_result.set_index('Sec Name')['Total FTE'])
                st.info(f"Original Total FTE: {original_fte:.2f}")
                st.info(f"Generated FTE: ${generated_fte:,.2f}")
            else:
                st.warning("Course not found.")
    else:
        st.warning("This feature will run when 'Sec Name' is present in the FTE dataset.")
