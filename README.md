# 使用說明

#### 測試前設定：

1. 下載Selenium Standalone Server (請參考<a href="http://www.seleniumhq.org/download/">Selenium Downloads</a>)

2. 下載瀏覽器WebDriver (請參考<a href="http://www.seleniumhq.org/download/">Selenium Downloads</a>)

3. 下載Web_Auto.jar及Web_TestScript.xlsm

#### 測試腳本建立說明：

1. 於C:\建立TUTK_QA_TestTool資料夾 (C:\TUTK_QA_TestTool)

2. TUTK_QA_TestTool中分別建立TestTool資料夾與TestReport資料夾

3. 將Web_TestScript.xlsm放至TestTool資料夾 (C:\TUTK_QA_TestTool\TestTool\Web_TestScript.xlsm)(檔名及副檔名請勿更改)

4. 開啟Web_TestScript.xlsm並允許啟動巨集 (已建立Web_Infor、ExpectResult及說明工作表)

5. Web_Infor工作表輸入Browser、BrowserDriverPath、TestURL、待測試腳本(以_TestScript結尾的工作表)、Web_Auto.jar路徑及Selenium Standalone Server.jar路徑，範例如下圖：

![image](https://github.com/Gilleschen/Android_invoke_excel/blob/master/picture/app_device_info_example.PNG)

6. 建立腳本(建立案列Case)：新增一工作表，工作表名稱須以_TestScript為結尾 (e.g. Login_TestScript)，目前支援指令如下: (大小寫有分，使用方式請參考Web_TestScript.xlsm內說明工作表)

        CaseName=>測試案列名稱(各案列開始時第一個填寫項目，必填!!!)

        Byid_Click=>搜尋元件id並點擊元件

        Byid_Result=>搜尋元件id並比對ExpectResult內容

        Byid_SendKey=>搜尋元件id並輸入數值或字串

        Byid_Wait=>等待並搜尋元件id

        ByXpath_Click=>搜尋元件xpath並點擊元件

        ByXpath_Result=>搜尋元件xpath並比對ExpectResult內容

        ByXpath_SendKey=>搜尋元件xpath並輸入數值或字串

        ByXpath_Wait=>等待並搜尋元件xpath

        Launch=>開啟瀏覽器並啟動指定的URL網址

        Quit=>關閉瀏覽器及WebDriver

        ScreenShot=>螢幕截圖

        Sleep=>閒置n秒鐘
  
範例腳本如下圖：

![image](https://github.com/Gilleschen/APP_Vsaas_2.0_Android_invoke_excel_Result_try_catch/blob/master/picture/Testcase_example.PNG)
  
7. ExpectResult工作表輸入各測試案例的期望結果

        7.1 A欄第二列處往下填入案列名稱 (CaseName)
        
        7.2 與案列名稱同列處輸入期望結果
        
 ExpectResult範例如下圖：
 
 ![image](https://github.com/Gilleschen/APP_Vsaas_2.0_Android_invoke_excel_Result_try_catch/blob/master/picture/Result_example.PNG)

#### 測試腳本語法檢查：

1. 執行Web_TestScript.xlsm增益集工具進行語法與資訊檢查，如下圖：

![image](https://github.com/Gilleschen/Android_invoke_excel/blob/master/picture/Gain_set.PNG)

2. 各功能說明：

        2.1 檢查資訊：確認Web_Infor工作表所有欄位是否正確
        
        2.2 檢查案例語法：確認各案例結束後均執行Quit方法，可不強制要求
        
        2.3 檢查案例輸入值：確認所有命令及參數是否正確
        
        2.4 檢查期望結果：確認案例之期望字串是否列於ExpectResult工作表，當然非所有案列都需列ExpectResult
        
        2.5 執行腳本：開始執行指定的工作表腳本，建議執行腳本前請確認前4項功能無誤
        
        註：2.2、2.3及2.4功能僅檢查以_TestScript為結尾且未隱藏的工作表 

#### Excel 測試報告

1. 開啟C:\TUTK_QA_TestTool\TestReport\Web_TestReport.xlsm

2. 根據瀏覽器類型自動建立TestReport工作表，如下圖： (e.g. chrome_TestReport)

![image](https://github.com/Gilleschen/APP_Vsaas_2.0_Android_invoke_excel_Result_try_catch/blob/master/picture/Testreport_sheet_example.PNG)

範例測試結果如下圖：

![image](https://github.com/Gilleschen/APP_Vsaas_2.0_Android_invoke_excel_Result_try_catch/blob/master/picture/Testreport_example.PNG)


