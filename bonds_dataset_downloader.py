import os, os.path
import win32com.client
import pandas as pd
import xlwings as xw
import datetime as dt
import time

#-------------------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------------

def runKDSQuery(spreadsheet):
    if os.path.exists(spreadsheet):
        xl = win32com.client.DispatchEx("Excel.Application")
        wb = xl.Workbooks.Open(os.path.abspath(spreadsheet))
        xl.Application.Run("'" + spreadsheet + "'!Module1.run_query")
        wb.Close(True)
        xl.Application.Quit()  
        del xl
    else:
        print("Couldn't fine spreadsheet named " + spreadsheet)
        
def editQuery(newQuery,spreadsheet):
    app = xw.App(visible=False)
    wb = xw.Book(spreadsheet)
    inputSheet = wb.sheets['Input']
    inputSheet.range('Query').value = newQuery
    wb.save()
    wb.close()
    app.kill()

def pickupResults(spreadsheet,outCSV='out.csv',mode='w',outputSheet="Output",resultsLabel=""):
    xlsx    = pd.ExcelFile(spreadsheet)
    output  = pd.read_excel(xlsx,sheet_name=outputSheet)
    output['Label'] = resultsLabel
    if os.path.exists(outCSV) == False:
        output.to_csv(outCSV,mode='w',index=False)
    else:
        if mode == 'w':
            output.to_csv(outCSV,mode='w',index=False)
        elif mode == 'a':
            output.to_csv(outCSV,mode='a',index=False,header=False)
        else:
            print("unknown to_csv writing mode: " + mode)            

#-------------------------------------------------------------------------------------------------------------------------------------------------
# Type=timeSeriesBy,x=asofdate, NonBlank=yes/y=cpr1: cpr3: cpr6: cpr12: cpr24: cprlife: smm: origb: currb: factor: ocoupon: coupon: owac: wac: wam: age: olnsz: clnsz: aols: waols: onloans: cnloans: osato: csato: oltv: cltv: fico: spread: %CashWindow: %Majors: PurpPctpurchase: PurpPctrefi: ChannelPctBroker: ChannelPctCorr: ChannelPctRetail: OccPctinvestor: OccPctowner: PropUnitsPct2-4: StatePctAK: StatePctAL: StatePctAR: StatePctAZ: StatePctCA: StatePctCO: StatePctCT: StatePctDC: StatePctDE: StatePctFL: StatePctGA: StatePctGU: StatePctHI: StatePctIA: StatePctID: StatePctIL: StatePctIN: StatePctKS: StatePctKY: StatePctLA: StatePctMA: StatePctMD: StatePctME: StatePctMI: StatePctMN: StatePctMO: StatePctMS: StatePctMT: StatePctNC: StatePctND: StatePctNE: StatePctNH: StatePctNJ: StatePctNM: StatePctNV: StatePctNY: StatePctOH: StatePctOK: StatePctOR: StatePctPA: StatePctPR: StatePctRI: StatePctSC: StatePctSD: StatePctTN: StatePctTX: StatePctUT: StatePctVA: StatePctVI: StatePctVT: StatePctWA: StatePctWI: StatePctWV: StatePctWY: SellerPctAMRHT: SellerPctALS: SellerPctCAFULL: SellerPctCNTL: SellerPctCITIZ: SellerPct53: SellerPctFIR: SellerPctFRDOM: SellerPctGUILD: SellerPctCHASE: SellerPctLLSL: SellerPctMATRX: SellerPctNCM: SellerPctNATIONSTAR: SellerPctNRESM: SellerPctPNYMAC: SellerPctPILOSI: SellerPctQUICK: SellerPctREG: SellerPctRMSC: SellerPctUNSHFI: SellerPctWFHM, Agency=umbs, MortgageType=fix, Program=sf, umbs=yes, DateWindowPeriod_by=cont: range: 202001: 202010: 1m, CusipPN_by=list: grid: 3136BCJH4: 3136BCUZ1: 3136BCJE1: 3136BCWR7: 3136BCUU2: 3136BCVE7: 3137FWQE3: 3137F6QQ3: 3137FXMX3: 3137FTQB6: 3137FXJ58: 3136BCDU1: 3137AY4Y4: 3136B7C82: 3136BBWL2: 3137FVRM6: 3136ADA90: 3136AM4F3: 3137F2X68: 3137FL4T8: 3136AMEG0: 3137FTKE6, JointDistribution=DateWindowPeriod_by: CusipPN_by,
def bondsAttributesQuery(observationWindow,CUSIPs='3136BCJH4: 3136BCUZ1: 3136BCJE1: 3136BCWR7: 3136BCUU2: 3136BCVE7: 3137FWQE3: 3137F6QQ3: 3137FXMX3: 3137FTQB6: 3137FXJ58: 3136BCDU1: 3137AY4Y4: 3136B7C82: 3136BBWL2: 3137FVRM6: 3136ADA90: 3136AM4F3: 3137F2X68: 3137FL4T8: 3136AMEG0: 3137FTKE6'):

    
    print(observationWindow)
    return "Type=timeSeriesBy,x=asofdate, NonBlank=yes/y=cpr1: cpr3: cpr6: cpr12: cpr24: cprlife: smm: origb: currb: factor: ocoupon: coupon: owac: wac: wam: age: olnsz: clnsz: aols: waols: onloans: cnloans: osato: csato: oltv: cltv: fico: spread: %CashWindow: %Majors: PurpPctpurchase: PurpPctrefi: ChannelPctBroker: ChannelPctCorr: ChannelPctRetail: OccPctinvestor: OccPctowner: PropUnitsPct2-4: StatePctAK: StatePctAL: StatePctAR: StatePctAZ: StatePctCA: StatePctCO: StatePctCT: StatePctDC: StatePctDE: StatePctFL: StatePctGA: StatePctGU: StatePctHI: StatePctIA: StatePctID: StatePctIL: StatePctIN: StatePctKS: StatePctKY: StatePctLA: StatePctMA: StatePctMD: StatePctME: StatePctMI: StatePctMN: StatePctMO: StatePctMS: StatePctMT: StatePctNC: StatePctND: StatePctNE: StatePctNH: StatePctNJ: StatePctNM: StatePctNV: StatePctNY: StatePctOH: StatePctOK: StatePctOR: StatePctPA: StatePctPR: StatePctRI: StatePctSC: StatePctSD: StatePctTN: StatePctTX: StatePctUT: StatePctVA: StatePctVI: StatePctVT: StatePctWA: StatePctWI: StatePctWV: StatePctWY: SellerPctAMRHT: SellerPctALS: SellerPctCAFULL: SellerPctCNTL: SellerPctCITIZ: SellerPct53: SellerPctFIR: SellerPctFRDOM: SellerPctGUILD: SellerPctCHASE: SellerPctLLSL: SellerPctMATRX: SellerPctNCM: SellerPctNATIONSTAR: SellerPctNRESM: SellerPctPNYMAC: SellerPctPILOSI: SellerPctQUICK: SellerPctREG: SellerPctRMSC: SellerPctUNSHFI: SellerPctWFHM, Agency=umbs, MortgageType=fix, Program=sf, umbs=yes, DateWindowPeriod_by=cont: range: " + observationWindow + ": 1m, CusipPN_by=list: grid: " + CUSIPs + ", JointDistribution=DateWindowPeriod_by: CusipPN_by,"
#-------------------------------------------------------------------------------------------------------------------------------------------------

def fullQueryRun(queryBuilder,observationWindows,issueWindow,outCSV,program='30'):

    starttime = time.perf_counter()
    
    if os.path.exists(outCSV):
        print("Removing ",outCSV)
        os.remove(outCSV)
        
    for observation_window in observationWindows:
        print('--------------------------------------------------------------------------------------')
        print(dt.datetime.now().strftime("%A, %B %d, %Y, %I:%M%p"))
    
        query = queryBuilder(observation_window,issueWindow,program)
        print(query)
        
        editQuery(query,spreadsheet);
        print('                    :  query edited, running')
        runKDSQuery(spreadsheet)
        print('                    :  query finished, collecting')
        pickupResults(spreadsheet,outCSV,mode='a',resultsLabel=issueWindow)
        print('                    :  query results appended')
        print('--------------------------------------------------------------------------------------')
    
    endtime = time.perf_counter()
    print('elapsed time        : ',"%1.1f" % (endtime - starttime),'sec /',"%1.2f" % ((endtime - starttime)/60),'min')
    print('--------------------------------------------------------------------------------------')

#------------------------------------------------------------------------------------------------------
    
#------------------------------------------------------------------------------------------------------
#---------------------------------- MAIN --------------------------------------------------------------
#------------------------------------------------------------------------------------------------------    
    
if __name__ == "__main__":
    
    spreadsheet = "Z:/Python Scripts/cpr-cdr runners/pools dataset/KDS/kds_macro.xlsm"
    data_dir    = "Z:/Python Scripts/cpr-cdr runners/pools dataset/data"
    
    todaysDate = dt.datetime.now(); 
        
    issue_periods = list(pd.Series(pd.date_range('2010-01-01',todaysDate,freq='M'),name='Issue Dates').apply(lambda x: x.strftime('%Y%m')))
    
    observation_periods = list(map(lambda x: str(x) + '01: ' + str(x) + '12',range(2010,2021,1)))
#    observation_periods = ['202002:202010']
#    observation_periods.append('current: current')

    #------------------------------------------------------------------------------------------------------   
    program = 'jumbo30'
    # download pool attributes data
    for i in range(len(issue_periods)):
        if i<1e6:
            issPeriod = issue_periods[i]
            year_of_issue = issPeriod[0:4]
            this_observ_periods = list()
            for ob_period in observation_periods:
                ob_period_start = ob_period.split(":")[0]
                if ob_period_start != 'current':
                    ob_period_year = ob_period_start[0:4]
                    if int(year_of_issue) <= int(ob_period_year):
                        this_observ_periods.append(ob_period)
            #this_observ_periods.append(observation_periods[-1])
            #------------------------------------------------------------------------------------------------------
#            print(f'{this_observ_periods} | {issPeriod}')
            fullQueryRun(poolAttributesQuery,this_observ_periods,issPeriod,data_dir + "/pools attributes/pools_attributes_issued_" + issPeriod + ".csv",program)
            #------------------------------------------------------------------------------------------------------
    #------------------------------------------------------------------------------------------------------  
    # download pool geographical data
    for i in range(len(issue_periods)):
        if i<1e6:
            issPeriod = issue_periods[i]
            year_of_issue = issPeriod[0:4]
            this_observ_periods = list()
            for ob_period in observation_periods:
                ob_period_start = ob_period.split(":")[0]
                if ob_period_start != 'current':
                    ob_period_year = ob_period_start[0:4]
                    if int(year_of_issue) <= int(ob_period_year):
                        this_observ_periods.append(ob_period)
#            this_observ_periods.append(observation_periods[-1])
            #------------------------------------------------------------------------------------------------------
            fullQueryRun(geoPctQuery,this_observ_periods,issPeriod,data_dir + "/geo pct/pools_geo_pct_issued_" + issPeriod + ".csv",program)
            #------------------------------------------------------------------------------------------------------            
    #------------------------------------------------------------------------------------------------------
    # download pool seller pct data
    for i in range(len(issue_periods)):
        if i<1e6:
            issPeriod = issue_periods[i]
            year_of_issue = issPeriod[0:4]
            this_observ_periods = list()
            for ob_period in observation_periods:
                ob_period_start = ob_period.split(":")[0]
                if ob_period_start != 'current':
                    ob_period_year = ob_period_start[0:4]
                    if int(year_of_issue) <= int(ob_period_year):
                        this_observ_periods.append(ob_period)
#            this_observ_periods.append(observation_periods[-1])
            #------------------------------------------------------------------------------------------------------
            fullQueryRun(poolSellerPctQuery,this_observ_periods,issPeriod,data_dir + "/seller pct/pools_attributes_issued_" + issPeriod + ".csv",program)
            #------------------------------------------------------------------------------------------------------
    
    #------------------------------------------------------------------------------------------------------
