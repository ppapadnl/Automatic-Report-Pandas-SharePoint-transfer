from google.oauth2 import service_account
import pandas as pd
import pycountry
from dateutil.relativedelta import relativedelta
import time
from datetime import datetime
import os
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
import logging


def setup_logging():
    log_directory = os.path.join(os.getcwd(), 'logs')
    if not os.path.exists(log_directory):
        os.makedirs(log_directory)
    log_filename = datetime.now().strftime("%Y-%m-%d_%H-%M-%S") + '.log'
    log_filepath = os.path.join(log_directory, log_filename)

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s [%(levelname)s] - %(message)s',
        handlers=[
            logging.FileHandler(log_filepath),
            logging.StreamHandler()  # To also output to console
        ]
    )


setup_logging()


def get_sharepoint_context_using_user():
    sharepoint_url = 'https://Account.sharepoint.com/sites/Report'
    user_credentials = UserCredential("USERNAME", "PASSWORD")
    ctx = ClientContext(sharepoint_url).with_credentials(user_credentials)
    return ctx


def upload_to_sharepoint(path: str, sharepointFolder_url: str):
    try:
        ctx = get_sharepoint_context_using_user()
        target_folder = ctx.web.get_folder_by_server_relative_url(sharepointFolder_url)
        file_name = os.path.basename(path)
        with open(path, 'rb') as content_file:
            file_content = content_file.read()
            target_folder.upload_file(file_name, file_content).execute_query()
            logging.info(f"Successfully uploaded {file_name} to SharePoint")
    except Exception as e:
        logging.error(f"Failed to upload {path} to SharePoint: {e}")


desire_month = datetime.now() - relativedelta(months=1)
formatted_month = desire_month.strftime("%B")

credentials = service_account.Credentials.from_service_account_file(
    r"C:\Users\Pypower\Documents\PythonProjects\Report\Key.json")
local_file_path = r"C:\Users\Pypower\Documents\PythonProjects\Report\MOT2.xlsx"
service_trans = r"C:\Users\Pypower\Documents\PythonProjects\Report\Carrier_services.xlsx"
b_unit = r"C:\UsersPypower\Documents\PythonProjects\Report\B_UNIT.xlsx"

# TMS BQ
sqlBigMile = """
WITH numbered_shipments AS (
  SELECT
    a.idSHP,
    substr(a.ShipDate, 0, 10) as ShipDate,
    s.zipCode,
    s.city,
    s.isoCountry,
    CASE
    when sha.ZipCode = 'BT70 1LF' then 'BT70' else sha.ZipCode end as ZIP,
    CASE
    when sha.City = 'Senlis, Oise' then 'Senlis Cedex' when sha.City = 'SENLIS OISE' then 'Senlis Cedex'
    when sha.City = 'Olching-Geisselbullach' then 'Olching' else sha.City end as City_1,
    CASE
    WHEN sha.Country = 'IC' THEN 'ES'
    WHEN sha.Country = 'AN' THEN 'CW'
    WHEN sha.Country = 'XK' THEN 'KX' ELSE sha.Country END as Country,
    CASE
    WHEN a.Weight <= 0 AND shatw.Data IS NOT NULL THEN CAST(ROUND(CAST(shatw.Data AS FLOAT64), 2) AS STRING)
    ELSE CAST(ROUND(CAST(a.Weight AS FLOAT64), 2) AS STRING) END as Weight,
    ROUND(a.Volume, 3) as Volume,
    sha.Name,
    a.OrderNo,
    s.name,
    UPPER(a.CodeCAR) as CodeCAR,
    a.ExternalId,
    UPPER(a.CodeCSE) as CodeCSE,
    a.CodeSEN,
    shat.Data as ReceiverReference1,
    sha.Address1 as Receiver_Address1,
    shasen.Address1 as Sender_Address1,
    ROW_NUMBER() OVER (PARTITION BY a.OrderNo ORDER BY a.Weight DESC) as row_num
  FROM `bq.Shipments` a
  JOIN `bq.Senders` s ON a.CodeSEN = s.CodeSEN
  JOIN `bq.ShipmentAddresses` sha ON a.idSHP = sha.idSHP AND sha.addressType = 'RECEIVER'
  JOIN `bq.ShipmentAddresses` shasen ON a.idSHP = shasen.idSHP AND shasen.addressType = 'SENDER'
  LEFT JOIN `bq.ShipmentAttributes` shat ON a.idSHP = shat.idSHP AND shat.Attribute = 'ReceiverReference1'
  LEFT JOIN `bq.ShipmentAttributes` shatw ON a.idSHP = shatw.idSHP AND shatw.Attribute = 'OriginalWeight'
  WHERE 1=1
    AND a.Status = '10'
    AND EXTRACT(YEAR FROM TIMESTAMP(a.ShipDate)) = EXTRACT(YEAR FROM TIMESTAMP(date_sub(current_date, INTERVAL 1 MONTH)))
    AND EXTRACT(MONTH FROM TIMESTAMP(a.ShipDate)) = EXTRACT(MONTH FROM TIMESTAMP(date_sub(current_date, INTERVAL 1 MONTH)))
    AND a.EndOfDayID != 0
    AND (
         (UPPER(a.CodeCAR) <> 'ZTRP')
      OR (UPPER(a.CodeCAR) = 'ZTRP' AND a.CodeSEN = 'LELY.NLSBR01@rhenus.com' AND a.CodeCSE = 'pendel')
      OR (UPPER(a.CodeCAR) = 'ZTRP' AND a.CodeSEN = 'BEA.NLSBR01@rhenus.com' AND a.CodeCSE = 'dhlfreight')
      OR (UPPER(a.CodeCAR) = 'ZTRP' AND a.CodeSEN = 'ELS.NLSBR01@rhenus.com' AND a.CodeCSE in ('DHLGF', 'nunner', 'Topspeed'))
      OR (UPPER(a.CodeCAR) = 'ZTRP' AND a.CodeSEN = 'FUNKO.NLSBR01@rhenus.com' AND a.CodeCSE = 'nunner')
        )
    AND NOT(
         (UPPER(a.CodeSEN) = 'CUSTOMERX' AND UPPER(a.CodeCAR) IN ('STD.UPSREADY.COM', 'STD.TNTEL.COM'))
      OR (UPPER(a.CodeSEN) = 'CUSTOMERX' AND UPPER(a.CodeCAR) IN ('DHL.NL', 'STD.TNTEL.COM'))
      OR (UPPER(a.CodeSEN) = 'CUSTOMERX' AND UPPER(a.CodeCAR) IN ('STD.FEDEXWS.COM', 'STD.TNTEL.COM', 'STD.DHL.COM'))
      OR (UPPER(a.CodeSEN) = 'CUSTOMERX' AND UPPER(a.CodeCAR) IN ('DHL.NL', 'STD.TNTEL.COM'))
      OR (UPPER(a.CodeSEN) IN ('CUSTOMERX' ,'CRLBC.NLTLG01@RHENUS.COM') AND UPPER(a.CodeCAR) = 'RHENUSROAD.NL')
      OR (UPPER(a.CodeSEN) = 'CUSTOMERX' AND UPPER(a.CodeCAR) IN ('CIBLEX.COM', 'STD.DHL.COM'))
      OR (UPPER(a.CodeSEN) = 'CUSTOMERX' AND UPPER(a.CodeCAR) IN ('DHL.NL', 'CIBLEX.COM', 'STD.DHL.COM', 'STD.UPSREADY.COM', 'TOF.DE'))
      OR (UPPER(a.CodeSEN) = 'CUSTOMERX' AND EXTRACT(YEAR FROM TIMESTAMP(a.ShipDate)) > 2023 AND UPPER(a.CodeCAR) = 'STD.DHL.COM')
      OR (UPPER(a.CodeSEN) = 'CUSTOMERX' AND UPPER(a.CodeCAR) = 'RHENUSROAD.NL')
      OR (UPPER(a.CodeSEN) IN ('CUSTOMERX' ,'FLIRS.NLTLG01@RHENUS.COM', 'CUSTOMERX') AND UPPER(a.CodeCAR) = 'CUSTOMERX.NL')
      OR (UPPER(a.CodeSEN) IN ('CUSTOMERX' ,'CUSTOMERX') AND UPPER(a.CodeCAR) = 'STD.UPSREADY.COM')
      OR (UPPER(a.CodeSEN) = 'CUSTOMERX' AND UPPER(a.CodeCAR) = 'STD.DHL.COM')
      OR (UPPER(a.CodeSEN) = 'CUSTOMERX' AND UPPER(a.CodeCAR) IN ('DHL.NL', 'STD.DHL.COM', 'STD.UPSREADY.COM', 'STD.TNTEL.COM'))
    )
)
SELECT *
FROM numbered_shipments
WHERE row_num = 1
ORDER BY ShipDate
"""

df = pd.read_gbq(sqlBigMile, project_id='bq', credentials=credentials)

# ISO3 Transformation and drop new column
df['ISO3'] = df['Country'].apply(
    lambda x: pycountry.countries.get(alpha_2=x).alpha_3 if pycountry.countries.get(alpha_2=x) else None)

df['CodeCSE'] = df['CodeCSE'].astype(str)

# READ XL MOT
xlsx_df = pd.read_excel(local_file_path)
xlsx_df['SERVICE_LEVEL'] = xlsx_df['SERVICE_LEVEL'].astype(str)

# READ XL Carrier_services
xlsx_df2 = pd.read_excel(service_trans)
xlsx_df2['CodeCSE'] = xlsx_df2['CodeCSE'].astype(str)

# READ XL Business Unit
xlsx_df3 = pd.read_excel(b_unit)
# xlsx_df3['CodeCSE'] = xlsx_df3['CodeCSE'].astype(str)


# MERGING DFs
result_df = pd.merge(df, xlsx_df[['CARRIER_ID', 'SERVICE_LEVEL', 'COUNTRY', 'POSTCODE', 'MODE_OF_TRANSPORT']],
                     how='left', left_on=['CodeCAR', 'CodeCSE', 'ISO3'],
                     right_on=['CARRIER_ID', xlsx_df['SERVICE_LEVEL'].str.upper(), 'COUNTRY'])

# Add a check for 'POSTCODE' to include only rows where 'ZIP' == 'POSTCODE' or 'POSTCODE' isnull
# result_df = result_df.query('ZIP == POSTCODE or POSTCODE.isnull()')

result_df = result_df[['idSHP', 'ShipDate', 'zipCode', 'city', 'isoCountry', 'ZIP', 'City_1', 'Country',
                       'Weight', 'Volume', 'Name', 'OrderNo', 'name_1', 'CodeCAR', 'ExternalId', 'CodeCSE', 'CodeSEN',
                       'ReceiverReference1',
                       'ISO3', 'Receiver_Address1', 'SERVICE_LEVEL', 'Sender_Address1', 'MODE_OF_TRANSPORT']]

# MERGING result_df with xlsx_df2(Carrier_service)
result_df2 = pd.merge(result_df, xlsx_df2[['CodeCAR', 'CodeCSE', 'Description']],
                      how='left', left_on=['CodeCAR', 'CodeCSE'],
                      right_on=[xlsx_df2['CodeCAR'].str.upper(), xlsx_df2['CodeCSE'].str.upper()])

result_df2 = result_df2[['idSHP', 'ShipDate', 'zipCode', 'city', 'isoCountry', 'ZIP', 'City_1', 'Country',
                         'Weight', 'Volume', 'Name', 'OrderNo', 'name_1', 'CodeCAR', 'ExternalId', 'CodeCSE', 'CodeSEN',
                         'ReceiverReference1',
                         'Description', 'ISO3', 'Receiver_Address1', 'SERVICE_LEVEL', 'Sender_Address1',
                         'MODE_OF_TRANSPORT']]

# MERGING result_df2 with xlsx_df3(Carrier_service)
result_df3 = pd.merge(result_df2, xlsx_df3[['Sender_code', 'Business_unit']],
                      how='left', left_on=['CodeSEN'], right_on=['Sender_code'])

result_df3 = result_df3[['idSHP', 'ShipDate', 'Sender_Address1', 'zipCode', 'city', 'isoCountry', 'Receiver_Address1',
                         'ZIP', 'City_1', 'Country', 'Weight', 'MODE_OF_TRANSPORT', 'name_1', 'Business_unit',
                         'ExternalId',
                         'CodeCAR', 'OrderNo', 'Name', 'ReceiverReference1', 'Description', 'Volume']]

file_name = fr"C:\Users\Pypower\Documents\PythonProjects\Report\Report_{formatted_month}.xlsx"
writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
print(file_name)
result_df3.to_excel(writer, startrow=0, sheet_name='Sheet1', header=True, index=False)
workbook = writer.book
worksheet = writer.sheets['Sheet1']
writer.close()
logging.info(f"DataFrame saved to {file_name}")
file_name = fr'Report_{formatted_month}.xlsx'
file_path = os.path.join(os.getcwd(), file_name)
sharepoint_folder_url = '/sites/Bigmile/BigMile/2025'  # Relative URL of the SharePoint folder
upload_to_sharepoint(file_path, sharepoint_folder_url)
logging.info("Script completed")
print(result_df3)
print("Job Completed check Logs")
