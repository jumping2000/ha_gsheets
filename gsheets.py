import hassapi as hass
import gspread
import sys
import datetime
from datetime import datetime, timedelta

#
# App which save data on Google Sheet
#
# Args:
#
# Release Notes
#
# Version 1.0:
#   Initial Version
#
# https://gspread.readthedocs.io/en/latest/
# https://github.com/burnash/gspread
#
class GSheets(hass.Hass):

    def initialize(self):
        self.log("###########  GSheets started ########### ")
        runtime = datetime.now()
        interval = self.args.get("RunEverySec")
        onceDay = self.args.get("RunOnceDay")
        self.path = self.args.get("PathJSON")

        if onceDay:
            self.run_daily(self.publish_to_gs,"23:59:40")
        else:
            runtime = runtime + timedelta(seconds=int(interval))
            self.run_every(self.publish_to_gs,runtime,int(interval))

    def publish_to_gs(self,kwargs):
        try:
            gc = gspread.service_account(filename=self.path)
            spreadSheet = gc.open(self.args["SpreadSheetName"])
        except gspread.exceptions.SpreadsheetNotFound as ex:
            self.log("Errore accesso al foglio Google {}".format(ex), level="ERROR")
            self.log(sys.exc_info())

        fmt = '%d-%m-%Y %H:%M:%S'
        d = datetime.now()
        d_string = d.strftime(fmt)
        
        uploadJson = self.args["upload"] #json string
        try:
            for option in uploadJson:
                if(self.checkSheetsCreated(spreadSheet ,option["sheetName"]) == False):
                    sheet = spreadSheet.add_worksheet(title=option["sheetName"], rows="100", cols="10")
                else:
                    sheet = spreadSheet.worksheet(option["sheetName"])
                names_list = []
                values_list = [] 
                for nome, valore in option.items():
                    if "nameInSheet" in nome:
                        names_list.append(valore)
                    elif "entity" in nome and self.get_state(valore) is not None:
                        values_list.append(float(self.get_state(valore)))
                index = 3
                names_list.insert(0,"Data - Orario")
                values_list.insert(0,d_string)
                if not "unknown" in values_list:
                    if sheet.cell(index-1, 1).value != "Data - Orario":
                        rng = gspread.utils.rowcol_to_a1(index-1, 1)+":"+ gspread.utils.rowcol_to_a1(index-1, len(names_list))
                        sheet.update(rng, [names_list])
                        sheet.format(rng, {
                                        "backgroundColor": {
                                            "red": 0.9,
                                            "green": 0.9,
                                            "blue": 0.9
                                            },
                                        "textFormat": {
                                            "bold": True
                                                }
                        })
                    sheet.insert_row(values_list, index, value_input_option='USER_ENTERED')
        except gspread.exceptions.GSpreadException:
            self.log("Errore nella elaborazione e salvataggio dei dati", level="ERROR")
            self.log(sys.exc_info())
 
    #restituisce una lista di tutti i fogli nel documento
    def checkSheetsCreated(self,spreadSheet,sheetName):
        listWorkSheets = spreadSheet.worksheets()
        for workSheet in listWorkSheets:
            if(sheetName == workSheet.title):
                return True
        return False
