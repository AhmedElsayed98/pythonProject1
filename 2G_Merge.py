

import pandas as pd
import numpy as np
import pyodbc
import PySimpleGUI as sg
from pathlib import Path

import sys


# In[6]:


layout = [
    [sg.Text("2G Dump:", font=("Helvetica", 10, "bold"),), sg.Input(key="-IN-"), sg.FileBrowse(file_types=(("Access Files", "*.mdb*"),))],
    [sg.Text('Add multiple Site Codes by space-separation i.e 0014UP 0521SI', font=("Helvetica", 10, "bold"), text_color='black')],
    [sg.Text('Site Code(s): ', font=("Helvetica", 10, "bold"),), sg.InputText(key="-IN2-", size=20)], 
    [sg.Text('Add multiple BCF IDs with space-separation i.e 1441 151', font=("Helvetica", 10, "bold"), text_color='black')],
    [sg.Text('New BCF ID(s): ', font=("Helvetica", 10, "bold"),), sg.InputText(key="-IN3-", size=20)],
    [sg.Text("Output Folder", font=("Helvetica", 10, "bold"),), sg.Input(key="-OUT-"), sg.FolderBrowse()],
    [sg.Exit(), sg.Button("Export CSVs!")]
]

window = sg.Window("2G Merge!", layout)


# In[7]:


def mergeStudy(MDB, SiteCode, newBCF, output_folder):
    # set up some constants
    
    SQL = ["SELECT A_BTS.BSCId, A_BTS.BCFId, A_BTS.BTSId, A_BTS.allowIMSIAttachDetach, A_BTS.amhLowerLoadThreshold, A_BTS.amhMaxLoadOfTgtCell, A_BTS.amhTrhoGuardTime, A_BTS.amhUpperLoadThreshold, A_BTS.bsIdentityCodeBCC, A_BTS.bsIdentityCodeNCC, A_BTS.btsIsHopping, A_BTS.btsLoadThreshold, A_BTS.btsMeasAver, A_BTS.callReestablishmentAllowed, A_BTS.cellBarQualify, A_BTS.cellBarred, A_BTS.cellLoadForChannelSearch, A_BTS.cellNumberInBtsHw, A_BTS.cellReselectHysteresis, A_BTS.cellReselectOffset, A_BTS.cellReselectParamInd, A_BTS.cnThreshold, A_BTS.dedicatedGPRScapacity, A_BTS.defaultGPRScapacity, A_BTS.dfcaUnsyncModeMaList, A_BTS.directGPRSAccessBts, A_BTS.diversityUsed, A_BTS.dlNoiseLevel, A_BTS.dldcEnabled, A_BTS.downgradeGuardTimeHSCSD, A_BTS.downwardGuardInterval, A_BTS.drInUse, A_BTS.drMethod, A_BTS.dtmEnabled, A_BTS.dtxMode, A_BTS.earlySendingIndication, A_BTS.emergencyCallRestricted, A_BTS.extendedBcchEnabled, A_BTS.failureThreshold, A_BTS.fastReturnToLTE, A_BTS.fddQMin, A_BTS.fddQMinOffset, A_BTS.fddQOffset, A_BTS.fddRscpMin, A_BTS.forcedHrCiAverPeriod, A_BTS.forcedHrModeHysteresis, A_BTS.frequencyBandInUse, A_BTS.gsmPriority, A_BTS.interferenceAveragingProcessAverPeriod, A_BTS.maxNumberOfRepetition, A_BTS.maxNumberRetransmission, A_BTS.measurementBCCHAllocation, A_BTS.msMaxDistInCallSetup, A_BTS.msPriorityUsedInQueueing, A_BTS.nbrOfSlotsSpreadTrans, A_BTS.nbrTchForPrioritySubs, A_BTS.newEstabCausesSupport, A_BTS.noOfBlocksForAccessGrant, A_BTS.noOfMFramesBetweenPaging, A_BTS.penaltyTime, A_BTS.qSearchI, A_BTS.qSearchP, A_BTS.queuePriorityNonUrgentHo, A_BTS.queueingPriorityCall, A_BTS.queueingPriorityHandover, A_BTS.rachDropRxLevelThreshold, A_BTS.radioFrequencyColorCode, A_BTS.radioLinkTimeout, A_BTS.radioLinkTimeoutAmrHrUlIncreaseStep, A_BTS.radioLinkTimeoutAmrUlIncreaseStep, A_BTS.radioLinkTimeoutUlIncreaseStep, A_BTS.reselectionAlgorithmHysteresis, A_BTS.rxLevAccessMin, A_BTS.rxSigLevAdjust, A_BTS.smsCbUsed, A_BTS.stircEnabled, A_BTS.timeLimitCall, A_BTS.timeLimitHandover, A_BTS.timerPeriodicUpdateMs, A_BTS.wcdmaPriority, A_BTS.insiteGateway, A_BCF.name FROM A_BCF INNER JOIN A_BTS ON (A_BCF.BCFId = A_BTS.BCFId) AND (A_BCF.BSCId = A_BTS.BSCId)", 
       "SELECT A_BTS_AMR.BSCId, A_BTS_AMR.BCFId, A_BTS_AMR.BTSId, A_BTS_AMR.amrHoFrThrDlRxQual, A_BTS_AMR.amrHoFrThrUlRxQual, A_BTS_AMR.amrHoHrInHoThrDlRxQual, A_BTS_AMR.amrHoHrThrDlRxQual, A_BTS_AMR.amrHoHrThrUlRxQual, A_BTS_AMR.amrPocFrPcLThrDlRxQual, A_BTS_AMR.amrPocFrPcLThrUlRxQual, A_BTS_AMR.amrPocFrPcUThrDlRxQual, A_BTS_AMR.amrPocFrPcUThrUlRxQual, A_BTS_AMR.amrPocHrPcLThrDlRxQual, A_BTS_AMR.amrPocHrPcLThrUlRxQual, A_BTS_AMR.amrPocHrPcUThrDlRxQual, A_BTS_AMR.amrPocHrPcUThrUlRxQual, A_BTS_AMR.amrSegLoadDepTchRateLower, A_BTS_AMR.amrSegLoadDepTchRateUpper, A_BTS_AMR.btsSpLoadDepTchRateLower, A_BTS_AMR.btsSpLoadDepTchRateUpper, A_BTS_AMR.radioLinkTimeoutAmr, A_BCF.name FROM A_BCF INNER JOIN A_BTS_AMR ON (A_BCF.BCFId = A_BTS_AMR.BCFId) AND (A_BCF.BSCId = A_BTS_AMR.BSCId)", 
       "SELECT A_TRX.BSCId, A_TRX.BCFId, A_TRX.BTSId, A_TRX.TRXId, A_TRX.initialFrequency, A_TRX.preferredBcchMark, A_TRX.gprsEnabledTrx, A_TRX.channel0Type, A_TRX.channel1Type, A_TRX.channel2Type, A_TRX.channel3Type, A_TRX.channel4Type, A_TRX.channel5Type, A_TRX.channel6Type, A_TRX.channel7Type, A_TRX.name, A_BCF.name FROM A_TRX INNER JOIN A_BCF ON (A_TRX.BCFId = A_BCF.BCFId) AND (A_TRX.BSCId = A_BCF.BSCId)",
       "SELECT A_POC.BSCId, A_POC.BCFId, A_POC.BTSId, A_POC.POCId, A_POC.alpha, A_POC.bepPeriod, A_POC.bsTxPwrMax, A_POC.bsTxPwrMax1x00, A_POC.bsTxPwrMin, A_POC.enableAla, A_POC.gamma, A_POC.maxPwrCompensation, A_POC.minIntBetweenAla, A_POC.pcALDlWeighting, A_POC.pcALDlWindowSize, A_POC.pcALUlWeighting, A_POC.pcALUlWindowSize, A_POC.pcAQLDlWeighting, A_POC.pcAQLDlWindowSize, A_POC.pcAQLUlWeighting, A_POC.pcAQLUlWindowSize, A_POC.pcControlEnabled, A_POC.pcControlInterval, A_POC.pcIncrStepSize, A_POC.pcLTLevDlNx, A_POC.pcLTLevDlPx, A_POC.pcLTLevUlNx, A_POC.pcLTLevUlPx, A_POC.pcLTQual144Nx, A_POC.pcLTQual144Px, A_POC.pcLTQual144RxQual, A_POC.pcLTQualDlNx, A_POC.pcLTQualDlPx, A_POC.pcLTQualDlRxQual, A_POC.pcLTQualUlNx, A_POC.pcLTQualUlPx, A_POC.pcLTQualUlRxQual, A_POC.pcLowerThresholdDlRxQualDhr, A_POC.pcLowerThresholdUlRxQualDhr, A_POC.pcLowerThresholdsLevDLRxLevel, A_POC.pcLowerThresholsLevULRxLevel, A_POC.pcRedStepSize, A_POC.pcUTLevDlNx, A_POC.pcUTLevDlPx, A_POC.pcUTLevUlNx, A_POC.pcUTLevUlPx, A_POC.pcUTQualDlNx, A_POC.pcUTQualDlPx, A_POC.pcUTQualDlRxQual, A_POC.pcUTQualUlNx, A_POC.pcUTQualUlPx, A_POC.pcUTQualUlRxQual, A_POC.pcUpperThresholdDlRxQualDhr, A_POC.pcUpperThresholdUlRxQualDhr, A_POC.pcUpperThresholdsLevDLRxLevel, A_POC.pcUpperThresholdsLevULRxLevel, A_POC.powerDecrQualFactor, A_POC.powerLimitAla, A_POC.pwrDecrLimitBand0, A_POC.pwrDecrLimitBand1, A_POC.pwrDecrLimitBand2, A_POC.tAvgT, A_POC.tAvgW, A_POC.transmitPowerReduction, A_BCF.name FROM A_POC INNER JOIN A_BCF ON (A_POC.BCFId = A_BCF.BCFId) AND (A_POC.BSCId = A_BCF.BSCId)",
        "SELECT A_HOC.BSCId, A_HOC.BCFId, A_HOC.BTSId, A_HOC.HOCId, A_HOC.allAdjacentCellsAveraged, A_HOC.allInterfCellsAveraged, A_HOC.allUtranAdjAver, A_HOC.amhTrafficControlIUO, A_HOC.amhTrafficControlMCN, A_HOC.amhTrhoPbgtMargin, A_HOC.averagingWindowSizeAdjCell, A_HOC.enaFastAveCallSetup, A_HOC.enaFastAveHo, A_HOC.enaFastAvePc, A_HOC.enaHierCellHo, A_HOC.enaTchAssSuperIuo, A_HOC.enableInterFrtIuoHo, A_HOC.enableIntraHoDl, A_HOC.enableIntraHoUl, A_HOC.enableMsDistance, A_HOC.enablePowerBudgetHo, A_HOC.enableUmbrellaHo, A_HOC.failMoveThreshold, A_HOC.fddRepThr2, A_HOC.gsmPlmnPriorisation, A_HOC.hoAvaragingLevDLWeighting, A_HOC.hoAvaragingLevDlWindowSize, A_HOC.hoAveragingLevUlWeighting, A_HOC.hoAveragingLevUlWindowSize, A_HOC.hoAveragingQualDlWeighting, A_HOC.hoAveragingQualDlWindowSize, A_HOC.hoAveragingQualUlWeighting, A_HOC.hoAveragingQualUlWindowSize, A_HOC.hoPeriodPbgt, A_HOC.hoPeriodUmbrella, A_HOC.hoTLDlPx, A_HOC.hoTLDlRxLevel, A_HOC.hoTLUlNx, A_HOC.hoTLUlPx, A_HOC.hoTLUlRxLevel, A_HOC.hoTQDlNx, A_HOC.hoTQDlPx, A_HOC.hoTQDlRxQual, A_HOC.hoTQUlNx, A_HOC.hoTQUlPx, A_HOC.hoTQUlRxQual, A_HOC.hoThrInterferenceDlNx, A_HOC.hoThrInterferenceDlPx, A_HOC.hoThresholdsInterferenceDlRxLevel, A_HOC.hoThresholdsInterferenceULNx, A_HOC.hoThresholdsInterferenceULPx, A_HOC.hoThresholdsInterferenceULRxLevel, A_HOC.hoThresholdsLevDLNx, A_HOC.interSystemDa, A_HOC.intfCellAvgWindowSize, A_HOC.intfCellNbrOfZeroResults, A_HOC.intraHoLoRxLevLimAmrHr, A_HOC.intraHoLoRxQualLimAmr, A_HOC.intraHoUpRxLevLimAmrHr, A_HOC.minBsicDecodeTime, A_HOC.minIntBetweenHoReq, A_HOC.minIntBetweenUnsuccHoAttempt, A_HOC.minIntIuoHoReqBQ, A_HOC.minIntUnsuccIsHo, A_HOC.minIntUnsuccIuoHo, A_HOC.msDHoThrParamN8, A_HOC.msDistanceAveragingParamHreqave, A_HOC.msDistanceHoThresholdParamMsRangeMax, A_HOC.msDistanceHoThresholdParamP8, A_HOC.multiratRep, A_HOC.noOfZeroResUtran, A_HOC.nonBcchLayerAccessThr, A_HOC.nonBcchLayerExitThr_nx, A_HOC.nonBcchLayerExitThr_px, A_HOC.numberOfZeroResults, A_HOC.oscDemultiplexingUlRxLevMargin, A_HOC.oscDhrDemultiplexingRxQualityThr, A_HOC.oscDhrMultiplexingRxQualityThr, A_HOC.oscMultiplexingUlRxLevelThr, A_HOC.oscMultiplexingUlRxLevelWindow, A_HOC.qSearchC, A_HOC.rxLevel, A_HOC.superReuseBadCiThresholdCiRatio, A_HOC.superReuseBadCiThresholdNx, A_HOC.superReuseBadCiThresholdPx, A_HOC.superReuseEstMethod, A_HOC.superReuseGoodCiThresholdCiRatio, A_HOC.superReuseGoodCiThresholdNx, A_HOC.superReuseGoodCiThresholdPx, A_HOC.thrDlRxQualDhr, A_HOC.thrUlRxQualDhr, A_HOC.utranAveragingNumber, A_HOC.utranHoThScTpdc, A_HOC.wcdmaRanCellPenalty, A_BCF.name FROM A_HOC INNER JOIN A_BCF ON (A_HOC.BCFId = A_BCF.BCFId) AND (A_HOC.BSCId = A_BCF.BSCId)",
      "SELECT A_BTS_PLMNPERMITTED.BSCId, A_BTS_PLMNPERMITTED.BCFId, A_BTS_PLMNPERMITTED.BTSId, A_BTS_PLMNPERMITTED.listId, A_BTS_PLMNPERMITTED.plmnPermitted, A_BCF.name FROM A_BTS_PLMNPERMITTED INNER JOIN A_BCF ON (A_BTS_PLMNPERMITTED.BCFId = A_BCF.BCFId) AND (A_BTS_PLMNPERMITTED.BSCId = A_BCF.BSCId)",
      "SELECT A_ADCE.BSCId, A_ADCE.BCFId, A_ADCE.BTSId, A_ADCE.adjacentCellIdMCC, A_ADCE.adjacentCellIdMNC, A_ADCE.adjacentCellIdLac, A_ADCE.adjacentCellIdCI, A_ADCE.adjCellBsicNcc, A_ADCE.adjCellBsicBcc, A_ADCE.bcchFrequency, A_BSC.name, A_BCF.name FROM A_BCF INNER JOIN (A_ADCE INNER JOIN A_BSC ON A_ADCE.BSCId = A_BSC.BSCId) ON (A_BCF.BCFId = A_ADCE.BCFId) AND (A_BCF.BSCId = A_ADCE.BSCId)",
      "SELECT A_ADJW.BSCId, A_ADJW.BCFId, A_ADJW.BTSId, A_ADJW.mnc, A_ADJW.mcc, A_ADJW.rncId, A_ADJW.AdjwCId, A_ADJW.lac, A_ADJW.uarfcn, A_ADJW.scramblingCode, A_BSC.name, A_BCF.name FROM A_BCF INNER JOIN (A_ADJW INNER JOIN A_BSC ON A_ADJW.BSCId = A_BSC.BSCId) ON (A_BCF.BCFId = A_ADJW.BCFId) AND (A_BCF.BSCId = A_ADJW.BSCId)",
      "SELECT A_BTS_GPRS.BSCId, A_BTS_GPRS.BCFId, A_BTS_GPRS.BTSId, A_BTS_GPRS.adaptiveLaAlgorithm, A_BTS_GPRS.cs3Cs4Enabled, A_BTS_GPRS.egprsEnabled, A_BTS_GPRS.egprsInitMcsAckMode, A_BTS_GPRS.egprsInitMcsUnAckMode, A_BTS_GPRS.egprsMaxBlerAckMode, A_BTS_GPRS.egprsMaxBlerUnAckMode, A_BTS_GPRS.egprsMeanBepOffset8psk, A_BTS_GPRS.egprsMeanBepOffsetGmsk, A_BTS_GPRS.extendedCellGprsEdgeEnabled, A_BTS_GPRS.extendedCellLocationKeepPeriod, A_BTS_GPRS.gprsCapacityThroughputFactor, A_BTS_GPRS.gprsDlPcEnabled, A_BTS_GPRS.gprsEnabled, A_BTS_GPRS.gprsMsTxPwrMaxCCH1x00, A_BTS_GPRS.gprsMsTxpwrMaxCCH, A_BTS_GPRS.gprsNonBCCHRxlevLower, A_BTS_GPRS.gprsNonBCCHRxlevUpper, A_BTS_GPRS.gprsRxlevAccessMin, A_BTS_GPRS.inactEndTimeHour, A_BTS_GPRS.inactEndTimeMinute, A_BTS_GPRS.inactStartTimeHour, A_BTS_GPRS.inactStartTimeMinute, A_BTS_GPRS.inactWeekDays, A_BTS_GPRS.csAckDl, A_BTS_GPRS.csAckUl, A_BTS_GPRS.csExtAckDl, A_BTS_GPRS.csExtAckUl, A_BTS_GPRS.csExtUnackDl, A_BTS_GPRS.csExtUnackUl, A_BTS_GPRS.csUnackDl, A_BTS_GPRS.csUnackUl, A_BTS_GPRS.initMcsExtAckMode, A_BTS_GPRS.initMcsExtUnackMode, A_BCF.name FROM A_BCF INNER JOIN A_BTS_GPRS ON (A_BCF.BCFId = A_BTS_GPRS.BCFId) AND (A_BCF.BSCId = A_BTS_GPRS.BSCId)"]
           
    Files = ["BTS", "BTS_AMR", "BTS_GPRS", "POC", "HOC", "PLMNPERMITTED", "ADCE_DELELTE", "ADJW_DELETE", "BTS_GPRS"]

    conn_string = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'+
    r'DBQ={};'.format(MDB))
    
    #con = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};' + r'DBQ={{{}}};'.format(MDB.replace("\\", "\\\\")))
    con = pyodbc.connect(conn_string)
    # oldBCF = [int(input("Old BCFId1: ")), int(input("Old BCFId2: "))]
    # BSC = int(input("BSC Code: "))
    # newBCF = int(input("New BCF: "))
    # 681 691 392853

    SiteCode = SiteCode.split(" ")
    newBCF = newBCF.split(" ")


    i = 0
    for ele in SQL:
        data = pd.read_sql(ele,con)
        #if (i == 6) | (i == 7): 
            #data['BSC name']= data.iloc[:,-2]
        #print(i)
        data['BCF name']= data.iloc[:,-1]
        data.drop('name', axis =1 , inplace=True)

        All = pd.DataFrame()
        for j in range(len(SiteCode)):
            data['Site Code'] = data['BCF name'].str[-6:]
            BTS = data[data['Site Code'] == SiteCode[j]]
            BTS = BTS[BTS['BCFId'] != int(newBCF[j])]
            #print(BTS)
            if (i != 6) & (i != 7):
                BTS['BTSId'] = int(newBCF[j]) + BTS['BTSId'].astype(str).str[-1:].astype(int) - 1
                BTS['BCFId'] = int(newBCF[j])

            All = pd.concat((All, BTS), axis=0)
            if (i == 6) | (i == 7):
                All['$operation'] = 'delete'
                #All.drop('BCF name', axis = 1,  inplace=True)

            if i == 5:
                All['plmnPermitted'] = 1

        All.drop(['Site Code', 'BCF name'], axis = 1,  inplace=True)
        All.dropna(axis=1, how='all', inplace=True)

        outputfile = Path(output_folder) / f"{Files[i]}.csv"
        All.to_csv(outputfile, index=False)
        i = i + 1

    SQL = SQL[6:8]
    i = 0
    Files = ['ADCE', 'ADJW']
    #SiteCode = SiteCode.split(" ")
    #newBCF = newBCF.split(" ")

    for ele in SQL:

        data = pd.read_sql(ele,con)

        data['BCF name']= data.iloc[:,-1]
        data['BSC name']= data.iloc[:,-3]
        data.drop('name', axis =1 , inplace=True)

        All = pd.DataFrame()
        for j in range(len(SiteCode)):
            data['Site Code'] = data['BCF name'].str[-6:]
            BTS = data[data['Site Code'] == SiteCode[j]]
            BTS = BTS[BTS['BCFId'] != int(newBCF[j])]

            BTS['BTSId'] = int(newBCF[j]) + BTS['BTSId'].astype(str).str[-1:].astype(int) - 1
            BTS['BCFId'] = int(newBCF[j])
            if i == 0:
                BTS['MML Command'] = "ZEAC:SEG=" + BTS['BTSId'].astype(str) + "::MCC=" + BTS['adjacentCellIdMCC'].astype(str) + ",MNC=" + BTS['adjacentCellIdMNC'].astype(str) + ",LAC=" + BTS['adjacentCellIdLac'].astype(str) + ",CI=" + BTS['adjacentCellIdCI'].astype(str) + ":NCC=" + BTS['adjCellBsicNcc'].astype(str) + ",BCC=" + BTS['adjCellBsicBcc'].astype(str) + ",BCC=" + BTS['bcchFrequency'].astype(str) +":SYNC=N;"
            else:
                BTS['MML Command'] = "ZEAE:SEG=" + BTS['BTSId'].astype(str) + ":INDEX=:MNC=" + BTS['mnc'].astype(str) + ",MCC=" + BTS['mcc'].astype(str) + ",RNC=" + BTS['rncId'].astype(str) + ",CI=" + BTS['AdjwCId'].astype(str) + ":LAC=" + BTS['lac'].astype(str) + ",SAC=" + BTS['AdjwCId'].astype(str) + ",FREQ=" + BTS['uarfcn'].astype(str) +",SCC=" + BTS['scramblingCode'].astype(str) + ",:;"

            MML = BTS[['MML Command', 'BSC name']].reset_index(drop=True)
            All = pd.concat((All, MML), axis=0)

        outputfile = Path(output_folder) / f"{Files[i]}.xlsx"
        All.to_excel(outputfile, index=False)
        i = i + 1
    

    sg.popup_no_titlebar("Merge CSVs exported Successfully ! \nBy: Othman Mohamed - Nokia NPO Engineer")


# In[8]:


def is_valid_path(filepath):
    if filepath and Path(filepath).exists():
        return True
    sg.popup_error("Filepath not correct!")
    return False

    
while True:
    #try:
    sg.theme('LightGrey1')
    event, values = window.read()
    if event in (sg.WINDOW_CLOSED, "Exit"):
        break

    if event == "Export CSVs!":
        if (is_valid_path(values['-IN-'])):
            mergeStudy(
                MDB = values["-IN-"],
                SiteCode = values["-IN2-"],
                newBCF = values["-IN3-"],
                output_folder= values["-OUT-"],
            )
            
   # except KeyError:
       # sg.popup_no_titlebar("Missing")
    #except PermissionError:
       # sg.popup_no_titlebar("Permission Denied, Kindly close the file!")
   # except FileNotFoundError:
       # sg.popup_no_titlebar("File Not Found!")
    #except:
       # sg.popup_no_titlebar("Undefined Error. Call the support") 
    
window.close()


# In[ ]:




