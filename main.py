import requests, argparse
from concurrent.futures import ThreadPoolExecutor
import pandas as pd

def prosess(org):
    response = requests.get(f"https://data.brreg.no/enhetsregisteret/api/enheter/{org}", headers={"accept": "application/json"})
    if response.status_code != 200:
        print(f"Org {org} gave statuscode {response.status_code}")
        return None
    enhets_data = response.json()

    response = requests.get(f"https://data.brreg.no/regnskapsregisteret/regnskap/{org}", headers={"accept": "application/json"})
    if response.status_code != 200:
        print(f"Org {org} gave statuscode {response.status_code}")
        return None
    regnskaps_data = response.json()[0]

    return [enhets_data["navn"], regnskaps_data["resultatregnskapResultat"]["aarsresultat"],regnskaps_data["egenkapitalGjeld"]["gjeldOversikt"]["sumGjeld"],regnskaps_data["egenkapitalGjeld"]["egenkapital"]["sumEgenkapital"]] # This line is probably to long
                                                                                                                                                                                                                                 # could probably use .append() instead
def main(Orgs,threads: int=8,formating: bool=True,output: str="Output.xlsx"):
    Data_order = ["Firma Navn","Årsresultat","Gjeld","Egenkapital"]

    print("Fetching Data")
    if threads <= 1:
        data = [prosess(org) for org in Orgs]
        
    else:
        with ThreadPoolExecutor(max_workers=threads) as executor:
            data = list(executor.map(prosess,Orgs))

    data = [val for val in data if val is not None]
    table = pd.DataFrame(data, columns=Data_order)
    writer = pd.ExcelWriter(output, engine="xlsxwriter")
    table.to_excel(writer, sheet_name="sheet1",index=False)

    workbook  = writer.book
    worksheet = writer.sheets['sheet1']
    if formating:
        print("formating data")
        format = workbook.add_format({'num_format': '0 kr'})
        worksheet.set_column(1,3,None,format)
        worksheet.autofit()
    workbook.close()
    print("Done")

if __name__ == "__main__":
        orgs = None
        parser = argparse.ArgumentParser(description="Program for fetching customer accounting data")
        parser.add_argument("--format",nargs="?",dest="formating",const=True,default=False)
        parser.add_argument("-f",nargs="?",type=str,dest="path",default="")
        parser.add_argument("-t",nargs="?",type=int,dest="threads",default=8)
        parser.add_argument("-d",nargs="*",type=int,dest="data",default=[])
        parser.add_argument("-o",nargs="?",type=str,dest="output",default="Output.xlsx")
        args = parser.parse_args()

        if args.path:
            with open(args.path.strip(),"r") as file:
                orgs = [val.strip()for val in list(file)]
        if args.data:
            orgs = list(args.data)
        if not orgs:
            raise Exception("Missing Requiered arguments -d or -f")

        main(orgs,args.threads,args.formating,args.output)