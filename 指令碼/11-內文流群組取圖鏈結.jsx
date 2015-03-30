/*取出文字流中群組物件內的圖鏈結 v1.00 */

/*================================== 變更日誌 ===========================================
2012-12-10  配合 InTools 6.0
================================== 變更日誌 ===========================================*/

var scriptName = "取出文字流中群組物件內的圖鏈結";

// 每個指令碼都要建立回傳陣列 陣列內容為字串
// [0] 要告訴爸爸的話    [1] 只允許True或False代表成功或失敗
var retAry = new Array(); // 回傳陣列
var s1 = ""; // 回傳給外部呼叫程式的文字訊息
var s2 = "False"; // 回傳給外部呼叫程式，本指令碼執行成功與否

var linkSrartSym = "<s-ItemLink-Group>";
var linkEndSym = "<e-ItemLink-Group>";
var symSegment = "//" //分段符號

//  從外部呼叫程式傳過來的參數要把值取出來
if( checkParas() == true ){ // 判斷外部呼叫程式是否給了程式所需參數的值   
    if( app.documents.length != 0 ){ // 要有開啟的文件
        var inDoc = app.activeDocument;
            // 務必加上錯誤擷取機制，以防指令碼無法將錯誤回傳
        try{
            /*        程式內容     */            
            extractGraphicsLinkInGroup();
            s1 = "取出文字流中群組物件內的圖鏈結，執行完畢!";
            s2 = "True"; // 執行成功
        }catch(e){
            s1 =  "取出文字流中群組物件內的圖鏈結，有錯誤發生!" + e.toString();
            s2 = "False"; // 執行失敗 
        }
    }else{
        s1 =  "沒有開啟的文件";
        s2 = "False"; // 執行失敗         
    }
}else{
    // 外部的呼叫程式沒有給予必要參數時
	s1 = " 沒有足夠必要的參數! ";
	s2 = "False";
}

// 最後把結果字串陣列回傳給外部呼叫程式
retAry[0] = s1;
retAry[1] = s2;
retAry; // 丟了

// ========================= Function ============================

function extractGraphicsLinkInGroup(){
    var win1 = new Window("palette",scriptName);
    var win1Label1 = staticTextMaker(win1);//進度訊息
    var win1Label2 = staticTextMaker(win1);//進度訊息
    var win1Label3 = staticTextMaker(win1);//錯誤訊息
    var progress1 = progressBarMaker(win1,inDoc.allPageItems.length);//進度條    
    var allObjItems = inDoc.allPageItems;
    for( var v1=inDoc.allPageItems.length-1 ; v1>=0 ; v1-- ){
        win1Label1.text = "識別群組物件中：["+(v1+1)+"]/["+inDoc.allPageItems.length+"]";
        progress1.value = inDoc.allPageItems.length-v1+1;
        if( inDoc.allPageItems[v1].constructor.name=="Group" ){ // 只針對群組物件
            var inGroup = inDoc.allPageItems[v1];
            if( inGroup.parent.constructor.name=="Character" && inGroup.allGraphics.length>0 ){ //群組物件是錨定在文字流中 且 有包含圖檔物件
                var parentChar = inGroup.parent;
                if( parentChar.insertionPoints.length==2 ){ // 必須有前後插入點
                    var aLinkStr = "";// 收集圖檔鏈結
                    for( var v2=inGroup.allGraphics.length-1 ; v2>=0 ; v2-- ){
                        var inGraphic = inGroup.allGraphics[v2];
                        aLinkStr += linkSrartSym + inGraphic.itemLink.name + linkEndSym;//是最後一個加結束符號
                        win1Label2.text = "圖形物件編號["+inGroup.id+"]：內文圖形["+(v2+1)+"]/["+inGroup.allGraphics.length+"]：長度["+aLinkStr.length+"]";
                        try{
                            inGraphic.remove();
                        }catch(e){
                            //刪除錯誤
                        }                        
                    }
                parentChar.insertionPoints[1].contents = aLinkStr;
                }
            }
        }
     }
//~     writeStrToFile( resultStr , "E:\\temp\\文流群組解\\result.txt" );
    progress1.parent.close();
}



// ====================== Sub Function ===========================

//檢查參數是否有輸入
function checkParas(){
    return true; // 不需參數
}


// 寫文字檔
function writeStrToFile( str , filePathStr ){
	var txtFile = new File (filePathStr);
	txtFile.open("w");
    if( str!=null ){
        txtFile.write ( str.toString() );
    }else{
        txtFile.write ( "NULL" );
    }
    txtFile.close();
}

//Create Platte Label
function staticTextMaker(parentWin)
{
    var txtLabel = parentWin.add("statictext");
    txtLabel.preferredSize = [500,30];
    txtLabel.text = "";
    return txtLabel;
}

//進度條
function progressBarMaker(parentWin,stopInt)
{
    var pbar = parentWin.add("progressbar",undefined,1,stopInt);
    pbar.preferredSize = [600,20];
    parentWin.show();
    return pbar;
}