﻿# functions for app.py

# Imports
import pandas as pd
from openpyxl.styles import Font, Border, Side, PatternFill
import options4 as opfour

def readfile():
    '''
    Generates the dataframe and then sorts it.

    Returns
    -------
    groups : dataframe
        the sorted dataframe of the file.

    '''
    try:
        # reads the deansDailyCsar.csv and unique_deansDailyCsar_FTE files in to a dataframe
        file_in = pd.read_csv('deanDailyCsar.csv')
        fte_file_in = pd.read_excel('unique_deansDailyCsar_FTE.xlsx')

        # merge prior dataframes
        # Extract Course Code from Sec Name if not already done
        if "Course Code" not in file_in.columns:
            file_in["Course Code"] = file_in["Sec Name"].str.extract(r"([A-Z]{3}-\d{3})")

        # Also create Course Code in credits_df
        if "Course Code" not in fte_file_in.columns:
            fte_file_in["Course Code"] = fte_file_in["Sec Name"].str.extract(r"([A-Z]{3}-\d{3})")

        # Merge only needed columns from credits_df
        merged_df = pd.merge(
            file_in,
            fte_file_in[["Course Code", "Contact Hours"]],
            how='left',
            on='Course Code'
        )

        merged_df["Contact Hours"] = pd.to_numeric(
        merged_df["Contact Hours"], errors='coerce')
        merged_df["FTE Count"] = pd.to_numeric(merged_df["FTE Count"], errors='coerce')

        # Calculate Total FTE
        merged_df["Total FTE"] = ((merged_df["Contact Hours"] * 16 *\
                                   merged_df["FTE Count"]) / 512).round(3)


        # sorts the dataframe by sec divisions, sec name
        # and sec faculty info and assigns it to groups
        groups = merged_df.sort_values(["Sec Divisions", "Sec Name", "Sec Faculty Info"])

        return groups

    except FileNotFoundError:
        groups = []
        print("File Missing!")
        return groups


def calc_enrollment(row):
        try:
            cap = float(row["Capacity"])
            fte = float(row["FTE Count"])

            if cap == 0:
                return "0%"

            percentage = (fte / cap) * 100
            return f"{percentage:.2f}%"

        except (ValueError, TypeError, ZeroDivisionError):
            return "N/A%"


def fte_by_div_raw(file_in, fte_tier, div_code):

    # Filter division
    div_code = div_code.upper()
    div_data = file_in[file_in['Sec Divisions'] == div_code].copy()

    if div_data.empty:
        return None, 0, 0

    # Create lookup for prefix/course ID → New Sector multiplier
    fte_lookup = {
        row['Prefix/Course ID']: row['New Sector']
        for _, row in fte_tier.iterrows()
        if pd.notna(row['Prefix/Course ID'])
    }

    # Extract course codes from Sec Name
    div_data['Course Code'] = div_data['Sec Name'].str.extract(r'([A-Z]+-\d+)')
    div_data = div_data.sort_values(['Course Code', 'Sec Name'])

    base_fte_value = 1926

    output_rows = []
    current_course = None
    course_total_fte = 0
    first_row = True

    grand_total_original_fte = 0
    grand_total_generated_fte = 0

    for _, row in div_data.iterrows():
        course = row['Course Code']
        sec = row['Sec Name'][:3] if pd.notna(row['Sec Name']) else ""

        new_sector_value = fte_lookup.get(sec, 0)
        total_fte = float(row['Total FTE']) if pd.notna(row['Total FTE']) else 0
        adjusted_fte = total_fte * (new_sector_value + base_fte_value)

        grand_total_original_fte += total_fte

        enrollment_per = ''
        if pd.notna(row['Capacity']) and pd.notna(row['FTE Count']) and float(row['Capacity']) > 0:
            enrollment_per = round((float(row['FTE Count']) / float(row['Capacity'])) * 100, 2)

        if course != current_course and current_course is not None:
            output_rows.append({
                'Division': '',
                'Course Code': 'Total',
                'Sec Name': '',
                'X Sec Delivery Method': '',
                'Meeting Times': '',
                'Capacity': '',
                'FTE Count': '',
                'Sec Faculty Info': '',
                'Total FTE': '',
                'Enrollment Per': '',
                'Generated FTE': course_total_fte
            })
            grand_total_generated_fte += course_total_fte
            course_total_fte = 0

        output_rows.append({
            'Division': div_code if first_row else '',
            'Course Code': course if course != current_course else '',
            'Sec Name': row['Sec Name'],
            'X Sec Delivery Method': row['X Sec Delivery Method'],
            'Meeting Times': row['Meeting Times'],
            'Capacity': row['Capacity'],
            'FTE Count': row['FTE Count'],
            'Sec Faculty Info': row['Sec Faculty Info'],
            'Total FTE': total_fte,
            'Enrollment Per': f"{enrollment_per}%" if enrollment_per != '' else '',
            'Generated FTE': adjusted_fte
        })

        course_total_fte += adjusted_fte
        current_course = course
        first_row = False

    # Add final course total
    if current_course is not None:
        output_rows.append({
            'Division': '',
            'Course Code': 'Total',
            'Sec Name': '',
            'X Sec Delivery Method': '',
            'Meeting Times': '',
            'Capacity': '',
            'FTE Count': '',
            'Sec Faculty Info': '',
            'Total FTE': '',
            'Enrollment Per': '',
            'Generated FTE': course_total_fte
        })
        grand_total_generated_fte += course_total_fte

    output_df = pd.DataFrame(output_rows)
    return output_df, grand_total_original_fte, grand_total_generated_fte


def format_fte_output(raw_df, original_fte_total, generated_fte_total):
    formatted_rows = []

    for _, row in raw_df.iterrows():
        formatted_row = row.copy()
        if isinstance(row['Generated FTE'], (float, int)):
            formatted_row['Generated FTE'] = "${:,.2f}".format(row['Generated FTE'])
        if isinstance(row['Total FTE'], (float, int)):
            formatted_row['Total FTE'] = "{:.3f}".format(row['Total FTE'])
        formatted_rows.append(formatted_row)

    df = pd.DataFrame(formatted_rows)
    df.loc[len(df.index)] = {
        'Division': '',
        'Course Code': 'DIVISION TOTAL',
        'Sec Name': '',
        'X Sec Delivery Method': '',
        'Meeting Times': '',
        'Capacity': '',
        'FTE Count': '',
        'Sec Faculty Info': '',
        'Total FTE': "{:.3f}".format(original_fte_total),
        'Enrollment Per': '',
        'Generated FTE': "${:,.2f}".format(generated_fte_total)
    }

    return df

def calculate_fte_by_course(df, fte_tier, course_code, base_fte=1926):

    course_code = course_code.upper()
    filtered = df[df['Course Code'] == course_code].copy()

    if filtered.empty:
        return None, 0, 0

    # Load FTE lookup
    fte_lookup = {
        row['Prefix/Course ID']: row['New Sector']
        for _, row in fte_tier.iterrows()
        if pd.notna(row['Prefix/Course ID'])
    }

    output_rows = []
    total_original_fte = 0
    total_generated_fte = 0

    for _, row in filtered.iterrows():
        sec_prefix = row['Sec Name'][:3]
        new_sector = fte_lookup.get(sec_prefix, 0)
        total_fte = float(row['Total FTE']) if pd.notna(row['Total FTE']) else 0
        generated_fte = total_fte * (new_sector + base_fte)
        total_original_fte += total_fte
        total_generated_fte += generated_fte

        enrollment_per = ''
        if pd.notna(row['Capacity']) and pd.notna(row['FTE Count']) and float(row['Capacity']) > 0:
            enrollment_per = round((float(row['FTE Count']) / float(row['Capacity'])) * 100, 2)

        output_rows.append({
            'Sec Name': row['Sec Name'],
            'X Sec Delivery Method': row['X Sec Delivery Method'],
            'Sec Faculty Info': row['Sec Faculty Info'],
            'Meeting Times': row['Meeting Times'],
            'Capacity': row['Capacity'],
            'FTE Count': row['FTE Count'],
            'Total FTE': total_fte,
            'Enrollment Per': f"{enrollment_per}%" if enrollment_per else '',
            'Generated FTE': generated_fte
        })

    # Add summary row
    output_rows.append({
        'Sec Name': 'COURSE TOTAL',
        'X Sec Delivery Method': '',
        'Sec Faculty Info': '',
        'Meeting Times': '',
        'Capacity': '',
        'FTE Count': '',
        'Total FTE': "{:.2f}".format(total_original_fte),
        'Enrollment Per': '',
        'Generated FTE': "${:,.2f}".format(total_generated_fte)
    })

    df_out = pd.DataFrame(output_rows)
    df_out['Total FTE'] = df_out['Total FTE'].apply(lambda x: "{:.2f}".format(x) if isinstance(x, float) else x)
    df_out['Generated FTE'] = df_out['Generated FTE'].apply(lambda x: "${:,.2f}".format(x) if isinstance(x, (float, int)) else x)

    return df_out, total_original_fte, total_generated_fte



def generate_faculty_fte_report(dean_df, fte_tier, faculty_name):
    """
    Generate an FTE report for a single faculty member.
    
    Parameters
    ----------
    dean_df : pd.DataFrame
        Merged dean dataset
    fte_tier : pd.DataFrame
        Tier multipliers
    faculty_name : str
        Full match of a faculty member from 'Sec Faculty Info'
    
    Returns
    -------
    pd.DataFrame
        Formatted FTE DataFrame
    float
        Total original FTE
    float
        Total generated FTE
    """
    frame = dean_df[dean_df["Sec Faculty Info"] == faculty_name].copy()
    frame = opfour.remove_duplicate_sections(frame)

    frame["Enrollment Per"] = opfour.calculate_enrollment_percentage(
        frame["FTE Count"], frame["Capacity"])

    frame = opfour.generate_fte(frame, fte_tier)

    total_original = frame["Total FTE"].sum()
    total_generated = frame["Generated FTE"].sum()

    frame["Total FTE"] = frame["Total FTE"].apply(lambda x: f"{x:.2f}")
    frame["Generated FTE"] = frame["Generated FTE"].apply(lambda x: f"${x:,.2f}")

    summary_row = pd.Series({
        "Sec Name": "TOTAL",
        "Total FTE": f"{total_original:.2f}",
        "Generated FTE": f"${total_generated:,.2f}"
    })

    return pd.concat([frame, pd.DataFrame([summary_row])], ignore_index=True), total_original, total_generated
