import os
import sys
import configparser
import pyodbc
import pandas as pd
import win32com.client as win32
from datetime import datetime
import file_management
from logger_updated_ipo_global_report import logger, error_email


class UpdatedIPOReport:
    def __init__(self):
        logger.info('-' * 100)
        self.config = configparser.ConfigParser()
        self.config.read('settings_update_ipo_report.ini')
        self.time_stamp = datetime.utcnow().strftime('%Y-%m-%d %H%M')
        self.results_folder = os.path.join(os.getcwd(), 'Results')
        self.result_file = os.path.join(self.results_folder, f"Filtered IPO Global Report {self.time_stamp}.xlsx")
        self.attachments_folder = os.path.join(os.getcwd(), 'Email Attachments')
        for folder in [self.results_folder, self.attachments_folder]:
            if not os.path.exists(folder):
                os.mkdir(folder)
        self.outlook = win32.Dispatch('Outlook.Application')
        self.df = self.latest_report_from_email()

    def latest_report_from_email(self, sub_folder: str = 'IPO Global Report') -> pd.DataFrame:
        """
        This method retrieves the latest report from my email, saves the report attached to the email,
        then returns a DataFrame created from the report.

        :param sub_folder: Name of the sub folder where the report is saved in my Inbox.
        :return: Pandas DataFrame created from the report attached to the latest email
        """
        # retrieving the attachment from the latest email in the IPO Global Report folder
        account = win32.Dispatch('Outlook.Application').Session.Accounts(1)
        inbox = self.outlook.GetNamespace('MAPI').Folders(account.DeliveryStore.DisplayName).Folders('Inbox')
        sub_folder = inbox.Folders(sub_folder)
        latest_message = sub_folder.Items.GetLast()
        latest_message_time = latest_message.LastModificationTime.strftime('%Y-%m-%d %H%M')
        logger.info(f"Latest report created at {latest_message_time}")
        if len(latest_message.Attachments) >= 1:
            # saving the attachment and returning DataFrame
            attachment = latest_message.Attachments[0]
            original_file_name = attachment.filename
            file_name, file_ext = os.path.splitext(original_file_name)
            new_file_name = file_name + ' ' + latest_message_time + file_ext
            new_file = os.path.join(self.attachments_folder, new_file_name)
            attachment.SaveAsFile(new_file)
            df = pd.read_csv(new_file)
            return df

    def filtering_report(self):
        """
        Filters the IPOs in the report for only IPOs where
        1). the expected listing date is in the future
        2). there is an IPO price, the IPO was priced in the last 7 days and it is not a blank check/SPAC IPO
        """
        for c in [c for c in self.df.columns if 'date' in c.lower()]:
            self.df[c] = pd.to_datetime(self.df[c].fillna(pd.NaT), errors='coerce')
        self.df = self.df.loc[
            (self.df['Listing_Date'] >= pd.to_datetime('today'))
            | (
                    (self.df['Pricing_Date'] >= (pd.to_datetime('today') - pd.offsets.DateOffset(days=7)))
                    & self.df['Price_per_Instrument'].notna()
                    & self.df['Blank_Check'].isna()
            )
        ]

    def return_db_connection(self, db_name: str):
        """
        Returns the connection object used to connect to the relevant database.
        Database connections are provided in an .ini file with the sections as database names.

        :param db_name: Database name provided in section of .ini file
        :return: pyodbc connection object
        """
        return pyodbc.connect(
            f"Driver={self.config.get(db_name, 'Driver')}"
            f"Server={self.config.get(db_name, 'Server')}"
            f"Database={self.config.get(db_name, 'Database')}"
            f"Trusted_Connection={self.config.get(db_name, 'Trusted_Connection')}",
            timeout=3)

    def tickers(self):
        """
        The tickers currently in the IPO Global Report are from Symbology. This method will run a query to retrieve
        the tickers collected in PEO-PIPE. If there are multiple tickers they will be concatenated (i.e. ABC, XYZ).
        The ticker(s) and exchange(s) returned from PEO-PIPE are then added to the data if the Symbology ticker is na.
        """
        logger.info("Getting ticker and exchange information from PEO-PIPE database")
        iconums = tuple(self.df['Iconum'].dropna().unique().tolist())
        query = self.config.get('query', 'peopipe') + ' ' + str(iconums)
        df_te = pd.read_sql_query(query, self.return_db_connection('termcond'))
        df_te['exchange'] = df_te['exchange'].str.strip()
        df_te.drop_duplicates(inplace=True)

        tickers = df_te[['iconum', 'ticker']]
        tickers = tickers.loc[tickers['ticker'].notna()]
        tickers.drop_duplicates(inplace=True)
        tickers = tickers.groupby('iconum')['ticker'].apply(', '.join).reset_index()
        
        exchanges = df_te[['iconum', 'exchange']]
        exchanges = exchanges.loc[exchanges['exchange'].notna()]
        exchanges = exchanges.loc[~exchanges['exchange'].str.contains('Not Traded')]
        exchanges.drop_duplicates(inplace=True)
        exchanges = exchanges.groupby('iconum')['exchange'].apply(', '.join).reset_index()

        df_comb = pd.merge(tickers, exchanges, how='outer', on='iconum')
        df_comb.rename(columns={col: col.title() for col in df_comb.columns}, inplace=True)
        self.df = pd.merge(self.df, df_comb, how='left', on='Iconum', suffixes=('_symb', '_peopipe'))
        # only replacing the ticker collected in PEO-PIPE if the Symbology ticker is null
        self.df['Ticker'] = self.df['Ticker_symb'].fillna(self.df['Ticker_peopipe'])
        self.df['Exchange'] = self.df['Exchange_symb'].fillna(self.df['Exchange_peopipe'])

    def format_data_frame(self):
        """
        Sorting the data frame so that the IPOs relevant for Symbology are at the top of the report,
        re-ordering the columns so that relevant information is in the first few columns,
        and formatting the date columns back to YYYY-MM-DD without time.
        """
        self.df.sort_values(by=['Listing_Date'], inplace=True)
        self.df.sort_values(by=['ISIN'], inplace=True, na_position='first')
        for c in [c for c in self.df.columns if 'date' in c.lower()]:
            self.df[c] = self.df[c].dt.strftime('%Y-%m-%d')
        good_cols = ['Company', 'Iconum', 'Filer_Type', 'Ticker', 'Exchange', 'FDS_CUSIP', 'ISIN', 'CUSIP',
                     'SEDOL', 'Price_per_Instrument', 'Listing_Date', 'Pricing_Date', 'Issue_Date', 'Trading_Date',
                     'Awareness_Date', 'Last_Updated', 'Security_Status', 'Security_Status_Date']
        bad_cols = ['Share_Type', 'Secondary_IPO', 'Currency', 'Domicile', 'Min_Price_per_Instrument',
                    'Max_Price_per_Instrument', 'Gross_Proceeds', 'Min_Gross_Proceeds', 'Max_Gross_Proceeds',
                    'Sponsored_Deal', 'Blank_Check', 'Document_Date', 'Deal_ID', 'Doc_ID', 'DAM_Doc_ID', 'Source_(URL)',
                    'Comments']
        self.df = self.df[good_cols]

    def save_results(self):
        self.df.to_excel(self.result_file, index=False, encoding='utf-8-sig')
        logger.info(f"File saved as {self.result_file}")
        
    def email_report(self):
        """
        Emails the report as an attachment.
        Email details like sender and recipients are provided in .ini file which is read by configparser.
        """
        mail = self.outlook.CreateItem(0)
        mail.To = self.config.get('email', 'to')
        mail.Sender = self.config.get('email', 'sender')
        mail.Subject = f"{self.config.get('email', 'subject')} {self.time_stamp}"
        mail.HTMLBody = self.config.get('email', 'body') + self.config.get('email', 'signature')
        if os.path.exists(self.result_file):
            mail.Attachments.Add(self.result_file)
        mail.Send()
        logger.info(f"Email sent to {self.config.get('email', 'to')}")
        

def main():
    try:
        up = UpdatedIPOReport()
        up.filtering_report()
        up.tickers()
        up.format_data_frame()
        up.save_results()
        up.email_report()
        file_management.main()
    except Exception as e:
        logger.error(e, exc_info=sys.exc_info())
        error_email(str(e))


if __name__ == '__main__':
    main()
