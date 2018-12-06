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
  

#### VBA 巨集

1. 點擊增益集，如下圖：

![image](https://github.com/Gilleschen/Web_Auto_Testing/blob/master/picture/functions.PNG)

2. 各功能說明：

* 執行腳本：執行指定的工作表腳本。

* 檢查資訊：檢查Web_Infor工作表欄位是否填寫。
        
* 檢查案例語法：確認各案例結束後均執行Quit方法。
        
* 檢查案例輸入值：確認所有指令及參數是否正確。
        
* 檢查期望結果：確認期望字串是否填入ExpectResult工作表。

* [腳本產生器](#scriptcreater)：透過VBA建立新腳本，也可手動建立工作表腳本。

#### VBA 功能異常排除

Step 1. 刪除增益集自訂工具列，如下圖：
        
![image](https://github.com/Gilleschen/Web_Auto_Testing/blob/master/picture/trobuleshotting.png)
        
Step 2. 存檔並關閉Web_TestScript.xlsm
        
Step 3. 重新開啟Web_TestScript.xlsm

# VBA 巨集使用說明

<a name="scriptcreater"/>

#### 腳本產生器說明 

Step 1. 點擊指令類型按鈕(藍框)，列出指令清單(綠框)

Step 2. 點選指令清單中的指令(綠框)後，點擊Add按鈕加入右側的腳本清單(紫框)

Step 3. 腳本完成後，點擊Create Case按鈕

![image](https://github.com/Gilleschen/Web_Auto_Testing/blob/master/picture/ScriptCreater.png)

##### 備註：

* Selenium Standalone Server:3.141.59

* Selenium Client Version:3.141.59

* Excel欄位若輸入純數字(e.g. 8888)，請轉換為文字格式，皆於數字前面加入單引號 (e.g. '8888)或執行增益集的檢查案例輸入值功能

* 僅支援Chrome, FireFox瀏覽器
