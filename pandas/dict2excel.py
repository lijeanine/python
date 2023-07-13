import pandas as pd
dict_data = [{
            "logType": "in",
            "requestMsg": "{\"version\":1,\"taskType\":\"10\",\"reported\":{\"CellData\":[{\"WipOrderNo\":\"860008804437\",\"SequenceNo\":\"000000\",\"SerialNo\":\"B22511NSD7A113B\",\"reqSys\":\"MesBus\",\"MainQuantity\":1100,\"DeviceID\":\"9900001103\",\"MainNGQuantity\":0,\"SourceSN\":\"B22511NCD7A103\",\"CompleteEmployeeNo\":\"80053098\",\"ActualCompleteDate\":\"2023-07-10 20:15:25\",\"ActualStartDate\":\"2023-07-09 13:43:15\",\"reqId\":\"SL01-1678377406806790144\",\"ProductNo\":\"\",\"OprSequenceNo\":\"0030\",\"SubNGQuantity\":0,\"TaskOrderNo\":\"DXTO202305260027\",\"ProgressType\":\"4\",\"Facility\":\"SL01\"}]},\"taskId\":\"1678377406806790144\"}",
            "responseMsg": "{\"msg\":\"success\",\"code\":\"0\",\"version\":1,\"taskId\":\"1678377406806790144\"}",
            "trackingNo": "1678377406806790144",
            "status": "SUCCESS",
            "type": "PROCESS_FEEDBACK",
            "sys": "IOT",
            "startDate": "2023-07-10 20:15:25",
            "endDate": "2023-07-10 20:15:25",
            "key": "10",
            "time": "2023-07-10 20:15:25.981",
            "level": "INFO",
            "contextName": "mom-service-me",
            "logTypeDesc": "上行",
            "statusDesc": "成功",
            "typeDesc": "生产进度反馈接口"
        },
        {
            "logType": "in",
            "requestMsg": "{\"subNgQuantity\":0,\"sourceSn\":\"B22511NCD7A103\",\"progressType\":4,\"reqSys\":\"MesBus\",\"deviceId\":\"9900001103\",\"actualCompleteDate\":\"2023-07-10 20:15:25\",\"reqId\":\"SL01-1678377406806790144\",\"wipOrderNo\":\"860008804437\",\"serialNo\":\"B22511NSD7A113B\",\"oprSequenceNo\":\"0030\",\"completeUserId\":70062,\"mainQuantity\":1100,\"startUserId\":70062,\"processTypeEnum\":\"PART\",\"completeEmployeeNo\":\"80053098\",\"actualStartDate\":\"2023-07-09 13:43:15\",\"taskOrderNo\":\"DXTO202305260027\",\"mainNgQuantity\":0,\"facility\":\"SL01\",\"productNo\":\"\",\"oprSerialNo\":\"000000\"}",
            "responseMsg": "Iot指令上报异步处理流程执行失败->任务单完工状态不匹配:订单状态为锁定",
            "trackingNo": "1678377406806790144",
            "status": "FAIL",
            "type": "PROCESS_FEEDBACK",
            "sys": "IOT",
            "startDate": "2023-07-10 20:15:25",
            "endDate": "2023-07-10 20:15:25",
            "key": "10-asynExecute",
            "time": "2023-07-10 20:15:25.998",
            "level": "INFO",
            "contextName": "mom-service-me",
            "logTypeDesc": "上行",
            "statusDesc": "失败",
            "typeDesc": "生产进度反馈接口"
        }]

def export_excel(dic_data):

    # if don`t need to change dic_data, we can omit try ... else 
    try:
        y = eval(dic_data[0]['responseMsg'])
    except:
        print(dic_data[0]['responseMsg'])
    else:
        dic_data[0].pop('responseMsg')
        dic_data[0] = dict(list(dic_data[0].items()) + list(y.items()))
        print(dic_data[0])
        
    # convert to DataFrame
    pf=pd.DataFrame(list(dic_data))
    order=['logType','trackingNo','status','type','sys','startDate','endDate','key','time','level','contextName','logTypeDesc','statusDesc','typeDesc','msg','code']
    pf=pf[order]
    print(pf)
    # name Excel
    file_path=pd.ExcelWriter('data_excel_0.xlsx')
    # replace space
    pf.fillna(' ', inplace=True)
    # output
    pf.to_excel(file_path, encoding='utf-8', index=False)
    # save excel
    file_path.save()
  
export_excel(dict_data)
