#!/usr/bin/env python
# coding: utf-8

# In[1]:


import os
import datetime
import pandas as pd
import numpy as np
import xlsxwriter
import streamlit as st
import zipfile


def edfinity_clean(edfinity_file):
    edf = pd.read_csv(edfinity_file)
    edf=edf[edf.columns.drop(list(edf.filter(regex='(Preview)')))]
    regularization_list=edf.drop(columns=['Last Name', 'First Name', 'Email/Username', 'ID', 'Course Name', 'Review of Prerequisites for Calculus I']).columns
    edf=edf.drop(columns='Review of Prerequisites for Calculus I')
    edf=edf.drop(columns='Edfinity Demo', errors = 'ignore')
    for column in regularization_list:
        edf[column]=(round(edf[column]/(edf[edf['First Name']=='Possible'][column].values), 2)>=.8).astype(int)
    edf=edf.drop('ID', axis=1)
    #edf=edf.drop(0, axis=0) #Drop Points Possible Row
    #assignment_list=['Edfinity '+item[0] for item in edf.columns[5:].str.split(' ')]
    #st.dataframe(edf)
    edf_summary=edf.sum(axis=0)
    edf=edf[edf.columns.drop(list(edf_summary[edf_summary==1].index))]
    return edf


# In[5]:

def aggregate_sheets(sheets_of_interest, convince_me_name, excel_tracker):
    long_sheet = pd.DataFrame()
    for sheet in sheets_of_interest:
        cache_sheet = pd.read_excel(excel_tracker, sheet_name = sheet).dropna(subset=['Student Name'])
        colMap = []

        for col in cache_sheet.columns:
            if col.rpartition('.')[0]:
                colName = col.rpartition('.')[0]
                inMap = col.rpartition('.')[0] in colMap
                lastIsNum = col.rpartition('.')[-1].isdigit()
                dupeCount = colMap.count(colName)

                if inMap and lastIsNum and (int(col.rpartition('.')[-1]) == dupeCount):
                    colMap.append(colName)
                    continue
            colMap.append(col)
        cache_sheet.columns=colMap
        cache_sheet = pd.melt(cache_sheet, id_vars = ['Student Name', 'Student ID'])
        cache_sheet['source'] = sheet
        long_sheet = pd.concat([long_sheet, cache_sheet], ignore_index=True)

    cm_sheet = pd.read_excel(excel_tracker, sheet_name = convince_me_name).dropna(subset=['Student Name'])
    cm_sheet.columns = ['Date', 'Student Name', 'Student ID', 'variable', 'value']
    cm_sheet = cm_sheet[['Student Name', 'Student ID', 'variable', 'value']]
    cm_sheet['source'] = convince_me_name

    long_sheet = pd.concat([long_sheet, cm_sheet], ignore_index=True)

    pwa_sheet = long_sheet[(((long_sheet.value==0) | (long_sheet.value==1)) & (long_sheet.source =='PWAs'))].dropna() #take out PWA indicators
    pwa_sheet = pwa_sheet.astype({
        'Student ID': 'int'
    })

    long_sheet = long_sheet[((long_sheet.value!=0) & (long_sheet.value!=1))].dropna() #take out PWA indicators

    long_sheet = long_sheet.dropna()

    long_sheet = long_sheet.astype(dtype= {"Student Name":"object",
                                           "Student ID":"int",
                                           "variable":"object",
                                           "value":"object",
                                           "source": "object",
                                          })
    long_sheet['mastery_points']=np.where(long_sheet.value=='Y', 1, 0)
    long_sheet['Category'] = np.where(long_sheet.variable.str.contains('*', regex=False), 'Supplementary', 'Core')
    return long_sheet, pwa_sheet, cm_sheet

# In[24]:

def student_emails(excel_tracker, sheet_name='Reference'):
    reference_sheet = pd.read_excel(excel_tracker, sheet_name = sheet_name).dropna(subset=['Student Name'])
    return reference_sheet


# In[7]:

def bad_edfinity_emails(edf, email_list):
    non_marian_emails = edf[~edf['Email/Username'].isin(email_list)]['Email/Username'].dropna()
    for email in non_marian_emails:
        replace_email=None
        temp_dict = {'nolanmac@outlook.com': 'nmacdonald727@marian.edu',
                     'hjminnis03@gmail.com': 'hminnis028@marian.edu',
                     'mnedohon369@marian.edu': np.nan,
                     'mjschelonka@gmail.com': 'mschelonka674@marian.edu'}
        replace_email = st.selectbox("Replace Non-Matching Email '"+email+"' with:",np.append(email_list, "Drop Student"),key=f"MyKey{email}")
        if replace_email is not None:
            if replace_email=='Drop Student':
                edf['Email/Username'] = edf['Email/Username'].replace(email, np.nan)
            else:
                edf['Email/Username'] = edf['Email/Username'].replace(email, replace_email)
    #
    return edf


# In[27]:

def edfinity_mapping(edf, attendance_sheet):
    mapped_edf = edf.dropna(subset=['Email/Username']).set_index('Email/Username').sum(axis=1).to_frame()
    mapped_edf=attendance_sheet[['Student ID', 'Preferred Email']].join(mapped_edf, on='Preferred Email').set_index('Preferred Email').dropna()
    #edfinity_mapping = pd.read_excel(uploaded_file, sheet_name="Edfinity Mapping")
    #mapped_edf = edfinity_mapping.merge(completed_edfinity, left_on = "Edfinity Email", right_on="Email/Username", how='outer')
    mapped_edf['Student ID'] = pd.to_numeric(mapped_edf['Student ID'],downcast='integer')
    return mapped_edf


# In[9]:

def set_mastery():
    core_mastery_targets = pd.DataFrame({
            'D' : [12, 0, 0, 0, 2, 19],
            'C' : [14, 0, 6, 0, 4, 21],
            'B': [16, 12, 8, 4, 6, 23],
            'A' : [18, 14, 10, 6, 8, 25]
        }, index=[2.88, #Core
                  3.18, #Core Continuing
                  1.85, #Supplementary
                  2.13, #Supp Continuing
                  .85, #PWA
                  -0.15]) #Edfinity

    core_mastery_targets["Mastery_Cat"]=["Core", "Core", "Supplementary", "Supplementary", "Professional Writing Assignments", "Edfinity"]#.reset_index(drop=True).melt()
    core_mastery_targets["Continuing_Flag"]=[False, True, False, True, False, False]
    core_mastery_targets["F"]=[0,0,0,0,0,0]

    #rejig table to insert as second sheet
    core_mastery_targets['index']=core_mastery_targets.Mastery_Cat+np.where(core_mastery_targets.Continuing_Flag==True, ' (CM)', '')
    mastery_table = core_mastery_targets.set_index('index').T.reset_index()[:4]
    store_column = mastery_table.pop('index')
    mastery_table.insert(len(mastery_table.columns), 'Grade', store_column)
    return mastery_table

# In[34]:


def workbook_writer(student_id, source_df, pwa_binary, edfinity_scores, mastery_table):
    #Filter source DF to create objective summary for a single student
    filtered_ls = source_df[~source_df['variable'].isin(['Date', 'Grade', 'Total', 'PWA Total'])] #originally long_sheet
    objective_summary = filtered_ls[filtered_ls['Student ID']==student_id].groupby(['Category', 'variable', 'source']).sum('mastery_points')[['mastery_points']]
    objective_summary = objective_summary.loc[(objective_summary!=0).all(axis=1)]
    objective_summary = (objective_summary.reset_index()
        .pivot(index=['Category','variable'], columns='source', values='mastery_points')
        .fillna(0)
    #    .astype({'PWAs': 'int','Quizzes': 'int','Tests': 'int'}, errors='ignore')
                            )
    choices = ['Continuing Mastery', 'Mastery', 'Not Mastered']
    objective_summary.loc[:,'Total'] = objective_summary.sum(axis=1)
    objective_summary['First']=np.where(objective_summary['Total']>=1, 'Y', '')
    objective_summary['Second']=np.where(objective_summary['Total']>=2, 'Y', '')
    objective_summary['Third']=np.where(objective_summary['Total']>=3, 'Y', '')

    #set filepath

    workbook = xlsxwriter.Workbook("reports/"+str(student_id)+".xlsx")
    workbook.formats[0].set_font_size(10)
    worksheet = workbook.add_worksheet(name="Targets")
    worksheet.set_default_row(12)
    reference = workbook.add_worksheet(name="Reference")

    #set formatting for headers and titles



    bold_format = workbook.add_format({'bold': True})

    header_format = workbook.add_format({'bold': False})
    header_format.set_bg_color('#031E51')
    header_format.set_font_color('white')

    subtotal_format = workbook.add_format({'bold': True})
    subtotal_format.set_align('right')
    subtotal_format.set_bg_color('#031E51')
    subtotal_format.set_font_color('white')

    long_subtotal_format = subtotal_format
    long_subtotal_format.set_text_wrap()


    mega_header_format = workbook.add_format({'bold': True})
    mega_header_format.set_bg_color('#031E51')
    mega_header_format.set_font_color('white')
    mega_header_format.set_font_size(14)
    mega_header_format.set_align('Center')
    mega_header_format_right = workbook.add_format({'bold': True})
    mega_header_format_right.set_bg_color('#031E51')
    mega_header_format_right.set_font_color('white')
    mega_header_format_right.set_font_size(14)
    mega_header_format_right.set_align('right')

    y_format = workbook.add_format()
    y_format.set_border(2)


    # Light red fill with dark red text.
    red_format = workbook.add_format({'bg_color':   '#FFC7CE',
                                   'font_color': '#9C0006'})

    # Light yellow fill with dark yellow text.
    yellow_format = workbook.add_format({'bg_color':   '#FFEB9C',
                                   'font_color': '#9C6500'})

    # Green fill with dark green text.
    green_format = workbook.add_format({'bg_color':   '#C6EFCE',
                                   'font_color': '#006100'})


    table_start = 12 #row index for first table row


    #Core Targets
    try:
        write_table = objective_summary.loc['Core']
    except:
        return
    core_targets_length = len(write_table)

    worksheet.write(table_start-1, 0, "",header_format)
    worksheet.write(table_start-1, 1, 'Mastery (M)', header_format)
    worksheet.write(table_start-1, 2, 'Continuing Mastery (CM)', header_format)
    worksheet.write(table_start-1, 3, '# of Ys', header_format)
    worksheet.write(table_start-1, 4, "",header_format)
    worksheet.write(table_start-1, 5, 'Enter Y', header_format)
    worksheet.write(table_start-1, 6, "",header_format)
    worksheet.write(table_start-1, 7, "",header_format)


    for i in range(core_targets_length):
        i_offset = i + table_start
        ##Mastery Column
        worksheet.write_formula(i_offset, 1, '=IF(D'+str(i_offset+1)+'>=2,"M"," ")')
        ##Continuing Mastery Column
        worksheet.write_formula(i_offset, 2, '=IF(D'+str(i_offset+1)+'=3,"CM"," ")')
        ##Total Y Column
        worksheet.write_formula(i_offset, 3, '=COUNTIF(E'+str(i_offset+1)+':G'+str(i_offset+1)+',"*Y*")')
        ##Fetch Y Columns
        worksheet.write(i_offset, 4, write_table['First'][i], y_format)
        worksheet.write(i_offset, 5, write_table['Second'][i], y_format)
        worksheet.write(i_offset, 6, write_table['Third'][i], y_format)
        worksheet.write(i_offset, 7, write_table.reset_index()['variable'][i], bold_format)
        worksheet.set_row(i_offset, None, None, {'level': 1})

    worksheet.write(i_offset+1, 0, 'Core Subtotal', subtotal_format)
    worksheet.write(i_offset+1, 1, '=COUNTIF(B'+str(table_start+1)+':B'+str(i_offset+1)+',"*M*")', header_format)
    worksheet.write(i_offset+1, 2, '=COUNTIF(C'+str(table_start+1)+':C'+str(i_offset+1)+',"*CM*")', header_format)
    worksheet.write(i_offset+1, 3, "", header_format)
    worksheet.write(i_offset+1, 4, "", header_format)
    worksheet.write(i_offset+1, 5, "", header_format)
    worksheet.write(i_offset+1, 6, "", header_format)
    worksheet.write(i_offset+1, 7, "", header_format)

    core_sub_row = i_offset+1
    try:
        write_table = objective_summary.loc['Supplementary']
    except:
        return
    supp_targets_length = len(write_table)

    for i in range(supp_targets_length):
        i_offset = i + table_start + core_targets_length+1
        ##Mastery Column
        worksheet.write_formula(i_offset, 1, '=IF(D'+str(i_offset+1)+'>=2,"M"," ")')
        ##Continuing Mastery Column
        worksheet.write_formula(i_offset, 2, '=IF(D'+str(i_offset+1)+'=3,"CM"," ")')
        ##Total Y Column
        worksheet.write_formula(i_offset, 3, '=COUNTIF(E'+str(i_offset+1)+':G'+str(i_offset+1)+',"*Y*")')
        ##Fetch Y Columns
        worksheet.write(i_offset, 4, write_table['First'][i], y_format)
        worksheet.write(i_offset, 5, write_table['Second'][i], y_format)
        worksheet.write(i_offset, 6, write_table['Third'][i], y_format)
        worksheet.write(i_offset, 7, write_table.reset_index()['variable'][i], bold_format)
        worksheet.set_row(i_offset, None, None, {'level': 1})

    worksheet.write(i_offset+1, 1, '=COUNTIF(B'+str(table_start+core_targets_length +2)+':B'+str(i_offset+1)+',"*M*")', header_format)
    worksheet.write(i_offset+1, 2, '=COUNTIF(C'+str(table_start+core_targets_length +2)+':C'+str(i_offset+1)+',"*CM*")', header_format)
    worksheet.write(i_offset+1, 0, 'Supplementary Subtotal', long_subtotal_format)
    worksheet.write(i_offset+1, 3, "", header_format)
    worksheet.write(i_offset+1, 4, "", header_format)
    worksheet.write(i_offset+1, 5, "", header_format)
    worksheet.write(i_offset+1, 6, "", header_format)
    worksheet.write(i_offset+1, 7, "", header_format)
    supp_sub_row = i_offset+1

    #Make it look nice with formatting

    worksheet.set_column('A:A', 24)
    worksheet.set_column('B:B', 10)
    worksheet.set_column('C:C', 30)
    worksheet.set_column('E:G', 2.5)

    worksheet.conditional_format('D'+str(table_start+1)+':D'+str(core_sub_row), {'type':     'cell',
                                            'criteria': '<=',
                                            'value':    1,
                                            'format':   red_format})

    worksheet.conditional_format('D'+str(core_sub_row+2)+':D'+str(i_offset+1), {'type':     'cell',
                                            'criteria': '<=',
                                            'value':    1,
                                            'format':   red_format})


    worksheet.conditional_format('D'+str(table_start+1)+':D'+str(core_sub_row), {'type':     'cell',
                                            'criteria': '=',
                                            'value':    2,
                                            'format':   yellow_format})

    worksheet.conditional_format('D'+str(core_sub_row+2)+':D'+str(i_offset+1), {'type':     'cell',
                                            'criteria': '=',
                                            'value':    2,
                                            'format':   yellow_format})

    worksheet.conditional_format('D'+str(table_start+1)+':D'+str(core_sub_row), {'type':     'cell',
                                            'criteria': '>=',
                                            'value':    3,
                                            'format':   green_format})

    worksheet.conditional_format('D'+str(core_sub_row+2)+':D'+str(i_offset+1), {'type':     'cell',
                                            'criteria': '>=',
                                            'value':    3,
                                            'format':   green_format})


    #### Begin Reference Sheet
    for j in range(len(mastery_table.columns)):
        k=0
        reference.write(k, j, mastery_table.columns[j])

        for k in range(len(mastery_table)):

            reference.write(k+1, j, mastery_table.loc[k][j])

    ##########Begin Summary Table

    worksheet.write(table_start-11, 0, '# mastered', bold_format)
    worksheet.write(table_start-11, 2, 'Category', bold_format)

    worksheet.write_formula(table_start-10, 0, '=B'+str(core_sub_row+1)) # (M) Core Learning Target
    worksheet.write_formula(table_start-10, 1, '=INDEX(Reference!$A$2:$G$5,MATCH(A3,Reference!$A$2:$A$5,1),7)')
    worksheet.write(table_start-10, 2, 'Core learning targets')


    worksheet.write_formula(table_start-9, 0, '=B'+str(supp_sub_row+1)) # (M) Supp Learning Target
    worksheet.write_formula(table_start-9, 1, '=INDEX(Reference!$A$2:$G$5,MATCH(A4,Reference!$C$2:$C$5,1),7)')
    worksheet.write(table_start-9, 2, 'Supplementary learning targets')


    worksheet.write(table_start-8, 0, '# continuing', bold_format)

    worksheet.write_formula(table_start-7, 0, '=C'+str(core_sub_row+1)) # (CM) Core Learning Target
    worksheet.write_formula(table_start-7, 1, '=INDEX(Reference!$A$2:$G$5,MATCH(A6,Reference!$B$2:$B$5,1),7)')
    worksheet.write(table_start-7, 2, 'Core learning targets')



    worksheet.write_formula(table_start-6, 0, '=C'+str(supp_sub_row+1)) # (CM) Supp Learning Target
    worksheet.write_formula(table_start-6, 1, '=INDEX(Reference!$A$2:$G$5,MATCH(A7,Reference!$D$2:$D$5,1),7)')
    worksheet.write(table_start-6, 2, 'Supplementary learning targets')


    ##PWA
    worksheet.write(table_start-4, 0, int((pwa_binary[pwa_binary['Student ID']==student_id].sum()['value'])))
    worksheet.write_formula(table_start-4, 1, '=INDEX(Reference!$A$2:$G$5,MATCH(A9,Reference!$E$2:$E$5,1),7)')
    worksheet.write(table_start-4, 2, 'Professional Writing Assignments')


    ##Edfinity
    worksheet.write(table_start-3, 0, int((edfinity_scores[edfinity_scores['Student ID']==student_id].sum()[0])))
    worksheet.write_formula(table_start-3, 1, '=INDEX(Reference!$A$2:$G$5,MATCH(A10,Reference!$F$2:$F$5,1),7)')
    worksheet.write(table_start-3, 2, 'Edfinity')



    ####Begin Header

    worksheet.write(table_start-12, 0, str(student_id), mega_header_format)
    worksheet.write(table_start-12, 1, "", mega_header_format)
    worksheet.write_formula(table_start-12, 2, '="GRADE: " & CHAR(LARGE(FILTER(CODE(B2:B11),B2:B11<>""),1))', mega_header_format)
    worksheet.write(table_start-12, 3, "", mega_header_format)
    worksheet.write(table_start-12, 4, "", mega_header_format)
    worksheet.write(table_start-12, 5, "", mega_header_format)
    worksheet.write(table_start-12, 6, "", mega_header_format)
    worksheet.write(table_start-12, 7, datetime.datetime.today().strftime('%Y-%m-%d'), mega_header_format_right)
    worksheet.set_row(table_start-12, 16)

    workbook.close()
    return

def zipdir(path, ziph):
            # ziph is zipfile handle
    for root, dirs, files in os.walk(path):
        for file in files:
            if ".xlsx" in file:
                ziph.write(os.path.join(root, file),
                             os.path.relpath(os.path.join(root, file),
                                             os.path.join(path, '..')))

# In[35]:


def midterm_targets_gen(tracker, sheet_name, midterm_date):
    learning_targets_timeline = pd.read_excel(tracker, sheet_name = sheet_name)
    filtered_targets = learning_targets_timeline[pd.to_datetime(learning_targets_timeline['Date'])<=pd.to_datetime(midterm_date)]
    agg_targets = pd.DataFrame(filtered_targets.T.drop(['Date', 'Assessment']).dropna(how='all', axis=0).sum(axis=1))
    agg_targets['Category'] = np.where(agg_targets.reset_index()['index'].str.contains('*', regex=False), 'Supplementary', 'Core')
    agg_targets['Category'].loc['PWA']='PWA'
    return agg_targets


def midterm_summary(source_df, midtern_targets):
    filtered_ls = source_df[~source_df['variable'].isin(['Date', 'Grade', 'Total', 'PWA Total'])] #originally long_sheet
    objective_summary = filtered_ls.groupby(['Student ID', 'Category', 'variable']).sum('mastery_points')[['mastery_points']]
    print(objective_summary.to_csv('obj.csv'))
    objective_summary = objective_summary.loc[(objective_summary!=0).all(axis=1)]
    objective_summary = (objective_summary.reset_index()
        .pivot(index=['Category','variable'], columns='source', values='mastery_points')
        .fillna(0)
        )
    st.dataframe(objective_summary)
    return

def main():
    st.title("Marian MBG Report Generation")

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




if __name__ == "__main__":
    main()
