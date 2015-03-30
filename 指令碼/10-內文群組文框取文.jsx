/*取出文字流中群組物件內的文字框文字 v1.1*/
/*====== 參數區( 開頭以 //$ 開始行且在程式內前30行者，均視為參數宣告，供外部主程式擷取 )============

========================*/

/*================================== 變更日誌 ===========================================
2012-04-25 群組物件完畢不刪，避免有東西未解出，或未成功解析刪除 extractTextFromTextFrameInGroup()
2012-12-10  配合 InTools 6.0
================================== 變更日誌 ===========================================*/

var scriptName = "取出文字流中群組物件內的文字框文字";

// 每個指令碼都要建立回傳陣列 陣列內容為字串
// [0] 要告訴爸爸的話    [1] 只允許True或False代表成功或失敗
var retAry = new Array(); // 回傳陣列
var s1 = ""; // 回傳給外部呼叫程式的文字訊息
var s2 = "False"; // 回傳給外部呼叫程式，本指令碼執行成功與否

var tfSrartSym = "<s-TextFrame-Group>";
var tfEndSym = "<e-TextFrame-Group>";
var symSegment = "//" //分段符號

//  從外部呼叫程式傳過來的參數要把值取出來
if( checkParas() == true ){ // 判斷外部呼叫程式是否給了程式所需參數的值   
    if( app.documents.length != 0 ){ // 要有開啟的文件
        var inDoc = app.activeDocument;
            // 務必加上錯誤擷取機制，以防指令碼無法將錯誤回傳
        try{
            /*        程式內容     */            
            extractTextFromTextFrameInGroup();
            s1 = "取出文字流中群組物件內的文字框文字，執行完畢!";
            s2 = "True"; // 執行成功
        }catch(e){
            s1 =  "取出文字流中群組物件內的文字框文字，有錯誤發生!" + e.toString();
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

function extractTextFromTextFrameInGroup(){
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
            if( inGroup.parent.constructor.name=="Character" && inGroup.textFrames.length>0 ){ //群組物件是錨定在文字流中 且 有包含文字框物件
                var parentChar = inGroup.parent;
                if( parentChar.insertionPoints.length==2 ){ // 必須有前後插入點
                    var aTextFrameContents = "";// 收集文框內容
                    for( var v2=inGroup.textFrames.length-1 ; v2>=0 ; v2-- ){
                        var inTextFrame = inGroup.textFrames[v2];
                        var inStory = inTextFrame.parentStory;//要以內文物為抽取標的
                        aTextFrameContents += tfSrartSym + inStory.contents + tfEndSym;//是最後一個加結束符號
                        win1Label2.text = "群組物件編號["+inGroup.id+"]：內文文框["+(v2+1)+"]/["+inGroup.textFrames.length+"]：長度["+aTextFrameContents.length+"]";
//~                         writeStrToFile( "===群組物件編號["+inGroup.id+"]：內文文框["+(v2+1)+"]/["+inGroup.textFrames.length+"]===\r\n"+aTextFrameContents , "E:\\log.txt" );
//~                         try{
//~                             inStory.remove();
//~                         }catch(e){
//~                             //刪除錯誤
//~                         }                        
                    }
                    //alert(aTextFrameContents);
                    try{
                        parentChar.insertionPoints[0].contents = aTextFrameContents;
                    }catch(e){
                        win1Label3.text = e.toString();//置入文字錯誤
                    }
                    //parentChar.insertionPoints[1].contents = aTextFrameContents;
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
	txtFile.open("a");
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