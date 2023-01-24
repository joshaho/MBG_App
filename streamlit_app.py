#!/usr/bin/env python
# coding: utf-8

# In[1]:


import grading_algorithms
import streamlit as st
from streamlit_option_menu import option_menu

# In[2] Excel App: 
def excel_app():
    st.title("Standards-Based Grading Report Generation")

    reports_ready=False
    if not os.path.exists('reports'):
        os.makedirs('reports')


    uploaded_file = st.file_uploader("Upload Grade Template", type = ['xlsx'])
    if uploaded_file is not None:
        excel_tracker = pd.ExcelFile(uploaded_file)
        sheet_select = excel_tracker.sheet_names
        if 'Cover' in sheet_select:
            sheet_select.remove('Cover')
            sheet_select.remove('Reference')
            sheet_select.remove('Learning Target Mapping')
        convince_me_name = st.selectbox("Select the sheet where Convince Me meetings are tracked", sheet_select)
        sheets_of_interest = st.multiselect("Select the grade sheets to be used (excluding Convince Me meetings)", sheet_select)
        generate_summary_flag = st.checkbox('Generate Class Summary')
        #midterm_flag = st.checkbox('Calculate Midterm Grades')
        #if midterm_flag:
    #        learning_targets_name = st.selectbox("Select the Learning Target Mapping sheet:", excel_tracker.sheet_names)
#            midterm_date = st.date_input("Select Midterm Cut-off")



    edfinity_file = st.file_uploader("Upload Edfinity Extract", type = ['csv'])
    if edfinity_file is not None:
        edf = edfinity_clean(edfinity_file)

    if ((edfinity_file is not None) and (uploaded_file is not None)):
        reference_sheet = student_emails(excel_tracker)
        email_list = reference_sheet['Preferred Email'].dropna().unique()
        edf = bad_edfinity_emails(edf, email_list)

        if st.button('Generate Reports'):

            st.write('Generating Reports...')
            long_sheet, pwa_sheet, cm_sheet = aggregate_sheets(sheets_of_interest, convince_me_name ,excel_tracker)
            mapped_edf = edfinity_mapping(edf, reference_sheet)
            st.download_button('[DEBUG] Download Raw Data', long_sheet.to_csv(), file_name='raw_grades.csv')
            mastery_table = set_mastery()
#            if midterm_flag:
#                midterm_targets = midterm_targets_gen(excel_tracker, learning_targets_name, midterm_date)
#                midterm_summary(long_sheet, midterm_targets)
            if generate_summary_flag:
                objective_summary = long_sheet.groupby(['Student ID','Category', 'variable']).sum('mastery_points')[['mastery_points']]
                objective_summary = objective_summary.loc[(objective_summary!=0).all(axis=1)]
                mastery_summary = objective_summary[objective_summary['mastery_points']>=2].reset_index().groupby(['Student ID', 'Category'])['variable'].nunique().reset_index()
                continuing_summary = objective_summary[objective_summary['mastery_points']>=3].reset_index().groupby(['Student ID', 'Category'])['variable'].nunique().reset_index()
                mastery_summary['Lvl'] = 'Mastery'
                continuing_summary['Lvl'] = 'Continuing Mastery'
                midframe = mastery_summary.append(continuing_summary)
                midframe['columns'] = midframe['Category'] + ' ' + midframe['Lvl']
                results_summary = (midframe.reset_index()
                    .pivot(index=['Student ID'], columns=['columns'], values='variable')
                    .fillna(0)
                    )
                frame_w_edf = mapped_edf.join(results_summary, on='Student ID')
                frame_w_edf.rename(columns  = {0:'Edfinity'}, inplace=True)
                frame_w_edf = frame_w_edf[['Student ID', 'Core Mastery', 'Core Continuing Mastery', 'Supplementary Mastery', 'Supplementary Continuing Mastery', 'Edfinity']]
                st.download_button('Download Summary', frame_w_edf.to_csv(), file_name='summary.csv')
            for id in pd.to_numeric(reference_sheet['Student ID'].dropna().unique(), downcast = 'integer'):
                workbook_writer(id, long_sheet, pwa_sheet, mapped_edf, mastery_table)
                reports_ready=True
    if reports_ready==True:
        with zipfile.ZipFile('reports/student_reports.zip', 'w', zipfile.ZIP_DEFLATED) as zipf:
            zipdir('reports/', zipf)

    if os.path.exists('reports/student_reports.zip') and reports_ready:
        with open('reports/student_reports.zip', 'rb') as f:
            download_button = st.download_button('Download Zip', f, file_name='student_reports.zip')  # Defaults to 'application/octet-stream'
        if download_button:
            os.remove('reports/student_reports.zip')
            report_ready=False
        clear_cache = st.button('Clear Cache')
        if clear_cache:
            os.remove('reports/student_reports.zip')
            report_ready=False
    return

# In[4] Define Sidebar:
def sidebar():
    with st.sidebar:
        choose = option_menu("Applications", ["Home", "Meeting Logger", "Assignment Input"],
                            icons=['house', 'kanban', 'book'],
                            menu_icon="app-indicator", default_index=0,
                            styles={
            "container": {"padding": "5!important", "background-color": "secondaryBackgroundColor"},
            "icon": {"color": "primaryColor", "font-size": "25px"}, 
            "nav-link": {"font-size": "16px", "text-align": "left", "margin":"0px", "--hover-color": "#eee"},
            "nav-link-selected": {"background-color": "backgroundColor"},
        }
        )
    return choose

# In[6] Login Function:

import streamlit as st

def check_password():
    """Returns `True` if the user had the correct password."""

    def password_entered():
        """Checks whether a password entered by the user is correct."""
        if st.session_state["password"] == st.secrets["password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # don't store password
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # First run, show input for password.
        st.text_input(
            "Password", type="password", on_change=password_entered, key="password"
        )
        return False
    elif not st.session_state["password_correct"]:
        # Password not correct, show input + error.
        st.text_input(
            "Password", type="password", on_change=password_entered, key="password"
        )
        st.error("😕 Password incorrect")
        return False
    else:
        # Password correct.
        return True


# In[5] Main Execution:
def main():
    menu_select = sidebar()
    if menu_select == "Home":
        #placeholder
        st.header("Home")
    elif menu_select == "Meeting Logger":
        st.header("Meeting Logger")
    elif menu_select == "Assignment Input":
        st.header("Assignment Input")


# In[3] Main Statement:
if __name__ == "__main__":
    if check_password():
        main()
