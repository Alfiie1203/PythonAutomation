import requests
import pandas as pd

# URL de la API para leer los datos
url_read = 'https://script.google.com/a/macros/vialto.com/s/AKfycbzvDr7aGqJF5X7gARXgbWT4hyVMASwAeDt4PhsIokznVD4rx6r6r8yERH0faJvf1PxAdA/exec?action=read'

# Realizar la solicitud GET deshabilitando la verificación SSL
response = requests.get(url_read, verify=False)  # <- Aquí se desactiva la verificación SSL
data = response.json()

# Crear un DataFrame de pandas con los datos recibidos
df = pd.DataFrame([data])

# Guardar el DataFrame en un archivo Excel

#excel_filename = 'datos_petition.xlsx'
#df.to_excel(excel_filename, index=False)

print(df)


#?action=add&data={"Name of the Petitioning Organization":"New Org","In Care Of Name (if any) last":"Smith","In Care Of Name (if any) first":"John","U.S. Street address":"123 Main St","City_Petitioner":"Los Angeles","State_Petitioner":"CA","Zip Code_Petitioner":"90001","Is this mailing address the same as the physical location of the sponsoring company or organization?":"Yes","Daytime Telephone Number":"555-555-5555","Email Address (if any)":"email@example.com","Website Address (if any)":"http://example.com","Does the petitioner employ 50 or more individuals in the United States?":"Yes","Are more than 50 percent of the petitioner's employees in H-1B, L-1A, or L-1B nonimmigrant status?":"No","The beneficiary will work as a:":"Engineer","startdate_proposed_employment":"2024-07-01","enddate_proposed_employment":"2025-07-01","Was the beneficiary of this petition in the United States during the last seven years?":"Yes","Family Name (Last Name)":"Doe","Given Name (First Name)":"Jane","Middle Name":"A","Street Number and Name or Po Box":"456 Elm St","City_Beneficiary":"San Francisco","Province_Beneficiary":"CA","PostalCode_Beneficiary":"94101","Country_Beneficiary":"USA","Date of Birth":"1990-01-01","Gender":"Female","City of birth":"San Francisco","State of birth":"CA","Country of birth":"USA","Country of Citizenship or Nationality":"USA","Number for the Blanket":"12345","Beneficiary's Wages Per Year":"80000","Beneficiary's Hours Per Week":"40","Job Title":"Software Engineer","Indicate the type of qualifying position the beneficiary was employed in while working for the qualifying foreign employer":"Engineer","foreign Employer Name":"Foreign Company","Street Address Mailing Address":"789 Maple St","City Mailing Address":"New York","Province Mailing Address":"NY","Postal Code Mailing Address":"10001","Country Mailing Address":"USA","Job Title_Act":"Software Engineer","Start Date":"2020-01-01","Wages Earned Per Year":"70000","Hours Worked Per Week":"40","With respect to the technology or technical data the petitioner will release or otherwise provide access to the beneficiary, the petitioner certifies that it has reviewed the Export Administration Regulations (EAR) and the International Traffic in Arms Regulations (ITAR) and has determined that:":"Compliant","Petitioner's or Authorized Signatory's Title":"Manager","Preparer's Family Name (Last Name)":"Smith","Preparer's Given Name (First Name)":"John","Preparer's Business or Organization Name":"Preparation Co.","Preparer's Daytime Telephone Number":"555-555-5556","Preparer's Email Address (if any)":"preparer@example.com"}
#?action=read
#
#https://script.google.com/a/macros/vialto.com/s/AKfycbzvDr7aGqJF5X7gARXgbWT4hyVMASwAeDt4PhsIokznVD4rx6r6r8yERH0faJvf1PxAdA/exec