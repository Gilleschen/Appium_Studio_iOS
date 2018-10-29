# 使用說明
#### 使用前設定：

* 安裝 Appium Studio (https://experitest.com/mobile-test-automation/appium-studio/)

* 下載 <a href="https://github.com/Gilleschen/Appium_Studio_iOS/blob/master/Appium_Studio_iOS.jar">Appium_Studio_iOS.jar</a>及<a href="https://github.com/Gilleschen/Appium_Studio_iOS/blob/master/TestScript_iOS.xlsm">TestScript_iOS.xlsm</a>

#### 測試腳本建立流程：

1. 於C:\建立TUTK_QA_TestTool資料夾 (C:\TUTK_QA_TestTool)

2. TUTK_QA_TestTool中分別建立TestTool資料夾與TestReport資料夾

3. 將TestScript_iOS.xlsm放至TestTool資料夾 (C:\TUTK_QA_TestTool\TestTool\TestScript_iOS.xlsm)(檔名及副檔名請勿更改)

4. 開啟TestScript_iOS.xlsm並允許啟動巨集 (已建立APP&Device、ExpectResult及說明工作表)

5. APP&Device工作表輸入APP bundle ID、測試裝置UDID、測試裝置OS版本、待測試腳本(以_TestScript結尾的工作表)、測試案例名稱、Appium_Studio_iOS.jar路徑及Reset APP，範例如下圖：

![image](https://github.com/Gilleschen/Appium_Studio_iOS/blob/master/img/APPandDevices.PNG)

6. 建立腳本：新增一工作表，工作表名稱須以_TestScript為結尾 (e.g. Login_TestScript)，目前支援指令如下: (大小寫有分，使用方式請參考TestScript.xlsm內說明工作表)

          CaseName=>測試案列名稱(各案列開始時第一個填寫項目，必填!!!)

          Back=>點擊手機返回鍵

          ByXpath_Click=>根據xpath搜尋元件並點擊元件

          ByXpath_LongPress=>根據xpath搜尋元件並長按元件

          ByXpath_VerifyText=>根據xpath搜尋元件並取得元件Text屬性之字串後，比對ExpectResult內容

          ByXpath_SendKey=>根據xpath搜尋元件並輸入數值或字串

          ByXpath_Clear=>根據xpath搜尋元件並清除數值或字串

          ByXpath_Wait=>根據xpath等待元件

          ByXpath_invisibility=>根據xpath搜尋元件並等待該元件消失

          ByXpath_Swipe=>根據xpath將元件A垂直移動到元件B位置,產生垂直滑動畫面

          ByXpath_Swipe_Horizontal=>垂直滑動/水平滑動n次

          Swipe=>根據x,y座標滑動畫面n次

          HideKeyboard=>關閉鍵盤

          Home=>點擊手機Home鍵

          LaunchAPP=>啟動APP&Device工作表指定的APP

          Orientation=>切換手機Landscape及Portrait模式

          Power=>點擊手機電源鍵

          QuitAPP=>關閉APP&Device工作表指定的Packageanme之Activity

          ResetAPP=>重置APP(清除APP暫存紀錄)並重新啟動APP

          ScreenShot=>螢幕截圖

          Sleep=>閒置APP n秒鐘
  
範例腳本如下圖：

![image](https://github.com/Gilleschen/APP_Vsaas_2.0_Android_invoke_excel_Result_try_catch/blob/master/picture/Testcase_example.PNG)
  
7. ExpectResult工作表輸入各測試案例的期望結果

   * A欄第二列處往下填入案列名稱 (CaseName)
        
   * 與案例名稱同列處輸入期望結果
        
 ExpectResult範例如下圖：
 
 ![image](https://github.com/Gilleschen/APP_Vsaas_2.0_Android_invoke_excel_Result_try_catch/blob/master/picture/Result_example.PNG)

#### 測試腳本語法檢查：

1. 執行TestScript_iOS.xlsm增益集工具進行語法與資訊檢查，如下圖：

![image](https://github.com/Gilleschen/Android_invoke_excel/blob/master/picture/Gain_set.PNG)

2. 各功能說明：

        2.1 檢查資訊：確認APP&Device工作表所有欄位是否正確
        
        2.2 檢查案例語法：確認各案例結束後均執行QuitAPP方法
        
        2.3 檢查案例輸入值：確認所有命令及參數是否正確
        
        2.4 檢查期望結果：確認案例之期望字串是否列於ExpectResult工作表，當然非所有案列都需列ExpectResult
        
        2.5 執行腳本：開始執行指定的工作表腳本，建議執行腳本前請確認前4項功能無誤
        
        註：2.2、2.3及2.4功能僅檢查以_TestScript為結尾且未隱藏的工作表 

3. 功能異常排除：

        3.1 移除增益集自訂工具列，如下圖：
        
      ![image](https://github.com/Gilleschen/Appium_Auto_Testing_Android/blob/master/picture/troubleshooting.png)
        
        3.2 存檔並關閉TestScript_iOS.xlsm
        
        3.3 重新開啟TestScript_iOS.xlsm
        
#### Appium Studio：

1. 啟動Appium Studio並設定iOS裝置 (請參考https://docs.experitest.com/display/TD/Appium+Studio)

2. 確認iOS裝置Ready，如下圖：
![image](https://github.com/Gilleschen/Appium_Studio_iOS/blob/master/img/device_setting.png)

#### Excel 測試報告

1. 開啟C:\TUTK_QA_TestTool\TestReport\TestReport_iOS.xlsm

2. 根據手機UDID自動建立TestReport工作表，如下圖： (e.g. abc123ABC123_TestReport)

![image](https://github.com/Gilleschen/APP_Vsaas_2.0_Android_invoke_excel_Result_try_catch/blob/master/picture/Testreport_sheet_example.PNG)

範例測試結果如下圖：

![image](https://github.com/Gilleschen/Web_Auto_Testing/blob/master/picture/TestResult.PNG)

#### 腳本產生器說明

點擊增益集中腳本產生器，如下圖：

![image](https://github.com/Gilleschen/Appium_Auto_Testing_Android/blob/master/picture/ScriptCreator.png)

1. 透過指令類型按鈕(籃框)，列出指令清單(綠框)

2. 點選指令清單中的指令(綠框)後，透過Add按鈕加入右側的腳本清單(紫框)

3. 腳本清單完成後，點擊Create Case按鈕

![image](https://github.com/Gilleschen/Appium_Auto_Testing_Android/blob/master/picture/ScriptCreator3.png)

#### 備註：

* Excel欄位若輸入純數字(e.g. 8888)，請轉換為文字格式，皆於數字前面加入單引號 (e.g. '8888)或執行增益集的檢查案例輸入值功能

* 固定Server Address = 127.0.0.1, 預設Server Port = 4725

* Appium NEW_COMMAND_TIMEOUT=120 second ;WebDriverWait timeout=30 second

* *目前不支援WiFi指令*


