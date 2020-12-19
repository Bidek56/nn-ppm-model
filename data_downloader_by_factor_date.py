import os, os.path
import win32com.client
import pandas as pd
import xlwings as xw
import datetime as dt
import time

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

# Type=PoolByPool,x=poolnumber, NonBlank=yes/y=cusip: prefix: spread: cpr1: cpr3: cpr6: cpr12: smm: DayCount: origb: currb: prevb: paydown: Prepay: factor: ocoupon: coupon: owac: wac: wam: age: aols: waols: cwals: owals: clnsz: olnsz: onloans: cnloans: pcnloans: ppnloans: osato: csato: oltv: cltv: FnBrk_OCLTV: FnBrk_CCLTV: fico: ODTI: CODTI: %CashWindow: %Majors: PurpPctpurchase: PurpPctrefi: ChannelPctBroker: ChannelPctCorr: ChannelPctRetail: OccPctowner: OccPct2ndHome: OccPctinvestor: PropUnitsPct2-4: Burnout: wac_min: wac_qtr1: wac_qtr3: wac_max: ofico_min: ofico_qtr1: ofico_qtr3: ofico_max: oltv_min: oltv_qtr1: oltv_qtr3: oltv_max: lnsz_min: lnsz_qtr1: lnsz_qtr3: lnsz_max: hpa3m: hpa1: hpa5: hpaLife: hpaPO3m: hpaPO1: hpaPO5: hpaPOLife, Agency=fn, MortgageType=fix, Program=30, DateWindowPeriod=202011, Issuance=201001: 202012,  poolType=bottom,
def poolAttributesQuery(observationWindow,issueWindow,program='30'):
    print(observationWindow + ' | ' + issueWindow)
    return "Type=PoolByPool,x=poolnumber, NonBlank=yes/y=cusip: prefix: spread: cpr1: cpr3: cpr6: cpr12: smm: DayCount: origb: currb: prevb: paydown: Prepay: factor: ocoupon: coupon: owac: wac: wam: age: aols: waols: cwals: owals: clnsz: olnsz: onloans: cnloans: pcnloans: ppnloans: osato: csato: oltv: cltv: FnBrk_OCLTV: FnBrk_CCLTV: fico: ODTI: CODTI: %CashWindow: %Majors: PurpPctpurchase: PurpPctrefi: ChannelPctBroker: ChannelPctCorr: ChannelPctRetail: OccPctowner: OccPct2ndHome: OccPctinvestor: PropUnitsPct2-4: Burnout: wac_min: wac_qtr1: wac_qtr3: wac_max: ofico_min: ofico_qtr1: ofico_qtr3: ofico_max: oltv_min: oltv_qtr1: oltv_qtr3: oltv_max: lnsz_min: lnsz_qtr1: lnsz_qtr3: lnsz_max: hpa3m: hpa1: hpa5: hpaLife: hpaPO3m: hpaPO1: hpaPO5: hpaPOLife, Agency=fn, MortgageType=fix, Program=" + program + ", DateWindowPeriod=" + observationWindow + ", Issuance=" + issueWindow + ", poolType=bottom,"

# Type=PoolByPool,x=poolnumber, NonBlank=yes/y=cusip: StatePoolPctAK: StatePoolPctAL: StatePoolPctAR: StatePoolPctAZ: StatePoolPctCA: StatePoolPctCO: StatePoolPctCT: StatePoolPctDC: StatePoolPctDE: StatePoolPctFL: StatePoolPctGA: StatePoolPctGU: StatePoolPctHI: StatePoolPctIA: StatePoolPctID: StatePoolPctIL: StatePoolPctIN: StatePoolPctKS: StatePoolPctKY: StatePoolPctLA: StatePoolPctMA: StatePoolPctMD: StatePoolPctME: StatePoolPctMI: StatePoolPctMN: StatePoolPctMO: StatePoolPctMS: StatePoolPctMT: StatePoolPctNC: StatePoolPctND: StatePoolPctNE: StatePoolPctNH: StatePoolPctNJ: StatePoolPctNM: StatePoolPctNV: StatePoolPctNY: StatePoolPctOH: StatePoolPctOK: StatePoolPctOR: StatePoolPctPA: StatePoolPctPR: StatePoolPctRI: StatePoolPctSC: StatePoolPctSD: StatePoolPctTN: StatePoolPctTX: StatePoolPctUT: StatePoolPctVA: StatePoolPctVI: StatePoolPctVT: StatePoolPctWA: StatePoolPctWI: StatePoolPctWV: StatePoolPctWY, Agency=fn, MortgageType=fix, Program=30, DateWindowPeriod=201001: current, Issuance=201001, poolType=bottom,
def geoPctQuery(observationWindow,issueWindow,program='30'):
    print(observationWindow + ' | ' + issueWindow)
    return "Type=PoolByPool,x=poolnumber, NonBlank=yes/y=cusip: StatePoolPctAK: StatePoolPctAL: StatePoolPctAR: StatePoolPctAZ: StatePoolPctCA: StatePoolPctCO: StatePoolPctCT: StatePoolPctDC: StatePoolPctDE: StatePoolPctFL: StatePoolPctGA: StatePoolPctGU: StatePoolPctHI: StatePoolPctIA: StatePoolPctID: StatePoolPctIL: StatePoolPctIN: StatePoolPctKS: StatePoolPctKY: StatePoolPctLA: StatePoolPctMA: StatePoolPctMD: StatePoolPctME: StatePoolPctMI: StatePoolPctMN: StatePoolPctMO: StatePoolPctMS: StatePoolPctMT: StatePoolPctNC: StatePoolPctND: StatePoolPctNE: StatePoolPctNH: StatePoolPctNJ: StatePoolPctNM: StatePoolPctNV: StatePoolPctNY: StatePoolPctOH: StatePoolPctOK: StatePoolPctOR: StatePoolPctPA: StatePoolPctPR: StatePoolPctRI: StatePoolPctSC: StatePoolPctSD: StatePoolPctTN: StatePoolPctTX: StatePoolPctUT: StatePoolPctVA: StatePoolPctVI: StatePoolPctVT: StatePoolPctWA: StatePoolPctWI: StatePoolPctWV: StatePoolPctWY, Agency=fn, MortgageType=fix, Program=" + program + ", DateWindowPeriod=" + observationWindow + ", Issuance=" + issueWindow + ", poolType=bottom,"

# Type=PoolByPool,x=poolnumber, NonBlank=yes/y=SellerPctAMRHT: SellerPctALS: SellerPctCAFULL: SellerPctCNTL: SellerPctCITIZ: SellerPct53: SellerPctFIR: SellerPctFRDOM: SellerPctGUILD: SellerPctCHASE: SellerPctLLSL: SellerPctMATRX: SellerPctNCM: SellerPctNATIONSTAR: SellerPctNRESM: SellerPctPNYMAC: SellerPctPILOSI: SellerPctQUICK: SellerPctREG: SellerPctRMSC: SellerPctUNSHFI: SellerPctWFHM: cusip: prefix, Agency=fn, MortgageType=fix, Program=30, DateWindowPeriod=current, Issuance=202001,  poolType=bottom,
def poolSellerPctQuery(observationWindow,issueWindow,program='30'):
    print(observationWindow + ' | ' + issueWindow)
    return "Type=PoolByPool,x=poolnumber, NonBlank=yes/y=SellerPctAMRHT: SellerPctALS: SellerPctCAFULL: SellerPctCNTL: SellerPctCITIZ: SellerPct53: SellerPctFIR: SellerPctFRDOM: SellerPctGUILD: SellerPctCHASE: SellerPctLLSL: SellerPctMATRX: SellerPctNCM: SellerPctNATIONSTAR: SellerPctNRESM: SellerPctPNYMAC: SellerPctPILOSI: SellerPctQUICK: SellerPctREG: SellerPctRMSC: SellerPctUNSHFI: SellerPctWFHM: cusip: prefix, Agency=fn, MortgageType=fix, Program=" + program + ", DateWindowPeriod=" + observationWindow + ", Issuance=" + issueWindow + ", poolType=bottom,"

def fullQueryRun(queryBuilder,query_running_spreadsheet,observation_window,issue_window,output_csv,program='30'):

    starttime = time.perf_counter()
    
    if os.path.exists(output_csv):
        print("Removing ",output_csv)
        os.remove(output_csv)
        
    print('--------------------------------------------------------------------------------------')
    print(dt.datetime.now().strftime("%A, %B %d, %Y, %I:%M%p"))

    query = queryBuilder(observation_window,issue_window,program)
    print(query)
    
    editQuery(query,query_running_spreadsheet);
    print('                    :  query edited, running')
    runKDSQuery(query_running_spreadsheet)
    print('                    :  query finished, collecting')
    pickupResults(query_running_spreadsheet,output_csv,resultsLabel="issue="+issue_window+" | observ="+observation_window)
    print('                    :  query results appended')
    print('--------------------------------------------------------------------------------------')
    
    endtime = time.perf_counter()
    print('elapsed time        : ',"%1.1f" % (endtime - starttime),'sec /',"%1.2f" % ((endtime - starttime)/60),'min')
    print('--------------------------------------------------------------------------------------')

    
if __name__ == "__main__":
    
    # needs to be full path or Excel throws an exception saying macro doesn't exist
    query_running_spreadsheet = "C:/Users/YuriTurygin/Desktop/NN-PPM/nn-ppm-model/KDS/kds_macro.xlsm"
    
    data_dir    = "../data"
    
    todaysDate = dt.datetime.now(); 
        
    issue_periods = ['201001:202012']
    
    observation_periods = ['202011']
    
    observation_periods = list()
    for y in range(2010,2021):
        for m in ['01','02','03','04','05','06','07','08','09','10','11','12']:
            observation_periods.append(str(y) + str(m))
    observation_periods.pop()
    

    #program = 'jumbo30'
    program = '30'

    # download pool attributes data
    for observation_window in observation_periods:
        if observation_window > '201701':
            for issue_window in issue_periods:
                fullQueryRun(poolAttributesQuery,query_running_spreadsheet,observation_window,issue_window,data_dir + "/pools attributes/pools_attributes_by_fct_date_" + observation_window + ".csv",program)

    # download pool geographical data
    for observation_window in observation_periods:
        for issue_window in issue_periods:
            fullQueryRun(geoPctQuery,query_running_spreadsheet,observation_window,issue_window,data_dir + "/geo pct/pools_geo_pct_by_fct_date_" + observation_window + ".csv",program)

    # download pool seller pct data
    for observation_window in observation_periods:
        for issue_window in issue_periods:
            fullQueryRun(poolSellerPctQuery,query_running_spreadsheet,observation_window,issue_window,data_dir + "/seller pct/pools_seller_pct_by_fct_date_" + observation_window + ".csv",program)
    





