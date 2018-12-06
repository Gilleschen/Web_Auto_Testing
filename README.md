# 使用說明

#### Framework
![image](https://github.com/Gilleschen/Web_Auto_Testing/blob/master/picture/framework.png)


#### 測試前設定：

* 下載Chorme, Firefox Third Party Browser Drivers(請參考<a href="http://www.seleniumhq.org/download/">Selenium Downloads</a>)

* 下載<a href="https://github.com/Gilleschen/Web_Auto_Testing/raw/master/Web_Auto.jar">Web_Auto.jar</a>及<a href="https://github.com/Gilleschen/Web_Auto_Testing/raw/master/Web_TestScrpit.xlsm">Web_TestScript.xlsm</a>至C:\TUTK_QA_TestTool\TestTool資料夾

* 建立C:\TUTK_QA_TestTool\TestReport資料夾


#### 腳本建立流程：

Step 1. 開啟Web_TestScript.xlsm並允許啟動巨集 (已建立Web_Infor、Web_InforData、ExpectResult及說明工作表)

Step 2. 建立腳本：新增一工作表，工作表名稱必需以_TestScript為結尾 (e.g. Login_TestScript)，目前支援指令如下: (大小寫有分，使用方式請參考Web_TestScript.xlsm內說明工作表)

* CaseName:測試案列名稱(各案列開始時第一個填寫項目，必填!!!)

* Byid_Click/ByXpath_Click:根據id/Xpath搜尋元件並點擊元件

* Byid_VerifyText/ByXpath_VerifyText:根據id/Xpath搜尋元件並比對ExpectResult內容

* Byid_SendKey/ByXpath_SendKey:根據id/Xpath搜尋元件並輸入數值或字串

* Byid_Wait/ByXpath_Wait:根據id/Xpath搜尋元件並等待元件出現
        
* Byid_invisibility/ByXpath_invisibility:根據id/Xpath搜尋元件並等待元件消失

* Launch:開啟瀏覽器並啟動指定的URL網址

* Quit:關閉瀏覽器及WebDriver

* ScreenShot:螢幕截圖

* Sleep:閒置n秒鐘
  
範例腳本如下圖：

![image](https://github.com/Gilleschen/Web_Auto_Testing/blob/master/picture/Script_example.PNG)

Step 3. 設定「期望字串」：點擊ExpectResult工作表，當使用Byid_VerifyText或ByXpath_VerifyText時，需在ExpectResult工作表填入期望字串。 (若測試案例不包含檢驗字串，則此步驟可省略)

* A欄第二列處往下填入案列名稱 (CaseName)
        
* 與案列名稱同列處輸入期望結果
        
 ExpectResult範例如下圖：
 
 ![image](https://github.com/Gilleschen/Android_invoke_excel/blob/master/picture/Result_example.PNG)
 
Step 4. 設定瀏覽器、測試網址資訊：點擊Web_Infor工作表，輸入Browser、BrowserDriverPath、TestURL、待測試腳本(以_TestScript結尾的工作表)，範例如下圖：

![image](https://github.com/Gilleschen/Web_Auto_Testing/blob/master/picture/web_infor.PNG)

Step 5. 點擊執行腳本，如下圖：
![image](https://github.com/Gilleschen/Web_Auto_Testing/blob/master/picture/RunScript.png)

#### Excel 測試報告

1. 開啟C:\TUTK_QA_TestTool\TestReport\Web_TestReport.xlsm

2. 根據瀏覽器類型自動建立TestReport工作表，如下圖： (e.g. chrome_TestReport)

![image](https://github.com/Gilleschen/Web_Auto_Testing/blob/master/picture/report.PNG)

範例測試結果如下圖：

![image](https://github.com/Gilleschen/Web_Auto_Testing/blob/master/picture/TestResult.PNG)
  

#### 測試腳本語法檢查：

1. 執行Web_TestScript.xlsm增益集工具進行語法與資訊檢查，如下圖：

![image](https://github.com/Gilleschen/Android_invoke_excel/blob/master/picture/Gain_set.PNG)

2. 各功能說明：

        2.1 檢查資訊：檢查Web_Infor工作表所有欄位
        
        2.2 檢查案例語法：確認各案例結束後均執行Quit方法
        
        2.3 檢查案例輸入值：檢查所有命令及參數
        
        2.4 檢查期望結果：確認各案例之期望字串是否列於ExpectResult工作表；若只想自動化操作Web，當然非所有案列都需列ExpectResult
        
        2.5 執行腳本：開始執行指定的工作表腳本，建議執行腳本前請確認前4項功能無誤
        
                2.5.1 點擊執行腳本後，自動啟動Selenium Hub (http://localhost:4444/)
        
                2.5.2 自動啟動Selenium Node (Port = 5555, maxInstances = 5)
        
        註：2.2、2.3及2.4功能僅檢查以_TestScript為結尾且未隱藏的工作表 

3. 功能異常排除：

        3.1 刪除增益集自訂工具列，如下圖：
        
      ![image](https://github.com/Gilleschen/Appium_Auto_Testing_Android/blob/master/picture/troubleshooting.png)
        
        3.2 存檔並關閉Web_TestScript.xlsm
        
        3.3 重新開啟Web_TestScript.xlsm



##### 備註：

* Selenium Client Version:3.8.1

* Excel欄位若輸入純數字(e.g. 8888)，請轉換為文字格式，皆於數字前面加入單引號 (e.g. '8888)或執行增益集的檢查案例輸入值功能

* 固定Selenium Node Port = 5555, maxInstances = 5

* 僅支援Chrome, FireFox, Internet Explorer瀏覽器

