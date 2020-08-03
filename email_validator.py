import pandas as pd
from validate_email import validate_email
import re
import xlsxwriter
import openpyxl

class Query_Emails:
    def setup_file(self, path):
        print('Creating Dataframe')
        global df
        global new_df
        df = pd.read_csv(path)
        new_df = df
        new_df['email'] = new_df['email'].fillna('').apply(str)
        new_df['regex'] = ''
        new_df['validation'] = ''
        new_df['stmp'] = ''
        new_df['verification'] = ''
        new_df['email status'] = ''
        new_df['domain'] = ''
        new_df['duplicates'] = ''

    def check_email_regex(self):
        global email_regex
        print('Checking Email Regex')
        email_regex = re.compile(r"""^[a-z0-9_.-]+@[0-9a-z.-]+\.[a-z.]{2,6}$""", re.X | re.I)
        new_df['regex'] = new_df.email.str.contains(email_regex, regex=True)

    def check_email_domain(self):
        print('Checking Email Domain')
        new_df[new_df['domain'].notnull()].copy()
        new_df['domain'] = new_df['email'].str.split('@').str[1]
        new_df.loc[pd.isnull(new_df['domain']) == True, 'domain'] = 'EMPTY'

    def check_email_duplicates(self):
        print('Checking Email Duplicates')
        new_df['duplicates'] = new_df['email'].duplicated()

    def check_email_validation(self):
        print('Checking Email Validation')
        new_df['validation'] = new_df['email'].apply(lambda x:validate_email(x)) # Validation evaluates if the mail conforms with the general mail regex
        new_df['stmp'] = new_df['email'].apply(lambda x: validate_email(x, check_mx=True)) # STMP Validation Check if the host has SMTP Server, return False for non-existing URLs
        new_df['verification'] = new_df['email'].apply(lambda x: validate_email(x, verify=True))  # Verification Check if the host has SMTP Server and the email really exists

class Analyse_Emails:
    def email_analysis_results(self):
        global status_validity_df, domain_validity_df, duplicate_dict
        global total_emails, valid, not_valid, domain_empty

        print('Analysing Email Statistics')
        total_emails = len(new_df['email'].index)
        duplicate_frequency = new_df['duplicates'].value_counts()
        duplicate_dict = duplicate_frequency.to_dict()

        print('Analysing Email Status Results')
        email_validity = []
        for index, row in new_df.iterrows():
            if row['regex'] == False:
                email_validity.append("Not Valid")
            elif row['regex'] and row['validation'] and row['stmp'] and row['verification'] == True:
                 email_validity.append("Valid")
            else:
                 email_validity.append("Unknown")
        new_df['email status'] = email_validity

        print('Analysing Email Status Results for Pie Chart')
        valid = new_df['email status'].str.contains(re.compile(r"""^Valid$""")).value_counts()[True]
        unknown = new_df['email status'].str.contains(re.compile(r"""^Unknown$""")).value_counts()[True]
        not_valid = new_df['email status'].str.contains(re.compile(r"""^Not Valid$""")).value_counts()[True]

        status_dict = {
            'Status': ['Valid', 'Unknown', 'Not Valid'],
            'Frequency': [valid, unknown, not_valid]
        }
        status_validity_df = pd.DataFrame(status_dict)

        print('Analysing Email Domain Results for Bar Chart')
        domain_frequency = new_df['domain'].value_counts()
        domain_dict = domain_frequency.to_dict()
        domain_validity_df = pd.DataFrame(list(domain_dict.items()), columns=['Domain', 'Frequency'])
        domain_empty = new_df['domain'].str.contains(re.compile(r"""^EMPTY$""")).value_counts()[True]

class Save_Emails:
    def csv(self, csv_path):
        print('Saving File as CSV')
        new_df.to_csv(csv_path, index=False)

    def excel(self, excel_path):
        print('Saving File as Excel')
        excel_file = excel_path

        #Creating Worksheets
        sheet_name_1 = 'Dashboard'
        sheet_name_2 = 'Emails'
        sheet_name_3 = 'Data'
        writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
        workbook = writer.book
        dashboard_worksheet = workbook.add_worksheet(sheet_name_1)
        new_df.to_excel(writer, sheet_name=sheet_name_2, index=False)
        status_validity_df.to_excel(writer, sheet_name=sheet_name_3, index=False)
        domain_validity_df.to_excel(writer, sheet_name=sheet_name_3, startcol = 2, index=False)

        #Setting 'Analysis' Background White
        cell_format = workbook.add_format()
        cell_format.set_pattern(1)
        cell_format.set_bg_color('white')
        dashboard_worksheet.set_column('A:Z', None, cell_format)

        #Email Stats
        merge_format = workbook.add_format({'align': 'left','valign': 'vcenter',})
        title_font_format = workbook.add_format({'align': 'center','valign': 'vcenter','bold': True, 'font_size': 18})
        dashboard_worksheet.merge_range('D3:H4', '', merge_format)
        dashboard_worksheet.write('D3:H4', 'Statistics', title_font_format)
        dashboard_worksheet.merge_range('D5:F5', 'Total Emails', merge_format)
        dashboard_worksheet.write_number('G5', total_emails)
        dashboard_worksheet.merge_range('D6:F6', 'Total Valid Emails', merge_format)
        dashboard_worksheet.write_number('G6', valid)
        dashboard_worksheet.merge_range('D7:F7', 'Total Not Valid Emails', merge_format)
        dashboard_worksheet.write_number('G7', not_valid)
        dashboard_worksheet.merge_range('D8:F8', 'Total Duplicates Emails', merge_format)
        dashboard_worksheet.write_number('G8', duplicate_dict[True])
        dashboard_worksheet.merge_range('D9:F9', 'Total Blank Emails', merge_format)
        dashboard_worksheet.write_number('G9', domain_empty)
        dashboard_worksheet.merge_range('B12:K13', '*If "Unknown" and the column "stmp" is empty = the domain is not valid or does not exist', merge_format)
        dashboard_worksheet.merge_range('B14:K15', '*If "Unknown" and the column "verification" is empty = either the name is not valid or the domain is unverifiable. It is not known if the email is valid or not.', merge_format)

        #Creating Pie Chart
        status_pie_chart = workbook.add_chart({'type': 'pie'})
        status_pie_chart.add_series({
            'name': 'Email Validity Percentage',
            'categories': '=Data!A2:A4',
            'values': '=Data!B2:B4',
        })
        status_pie_chart.set_chartarea({'border': {'none': True},})
        dashboard_worksheet.insert_chart('L3', status_pie_chart)

        # Creating Domain Bar Chart (Repeated domain chart)
        len_domain_names = len(domain_validity_df['Domain'])
        len_domain_frequency = len(domain_validity_df['Frequency'])

        domain_bar_chart = workbook.add_chart({'type': 'bar'})
        domain_bar_chart.add_series({
            'name': 'Email Domain Frequency',
            'categories': ['Data',1,2,len_domain_names,2],
            'values': ['Data',1,3,len_domain_frequency,3],
        })
        domain_bar_chart.set_chartarea({'border': {'none': True},})
        domain_bar_chart.set_legend({'none': True})
        domain_bar_chart.set_x_axis({'name': 'Frequency'})
        domain_bar_chart.set_y_axis({'name': 'Domain'})
        dashboard_worksheet.insert_chart('C17', domain_bar_chart)

        writer.save()

query = Query_Emails()
analyse = Analyse_Emails()
save = Save_Emails()

# ENTER INPUT CSV FILE LOCATION DIR
query.setup_file(r'')
query.check_email_regex()
query.check_email_domain()
query.check_email_duplicates()
query.check_email_validation()
analyse.email_analysis_results()
# ENTER SAVE CSV FILE LOCATION DIR
save.csv(r'')
# ENTER SAVE EXCEL FILE LOCATION DIR
save.excel(r'')


