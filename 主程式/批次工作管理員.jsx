/*批次工作管理員
    wenjen 2012-12-01 v1.0
*/
#target indesign;
#targetengine "session";

//Elément de publication
var jobsList; //ListBox
var grepList;
var scriptsList;
var msgText;//fenêtre de message

var grepFiles = getGrepFiles();//get grep files
var scriptFiles = getScriptFiles();//get script files
showWindow();

//============================= showWindow; ============================================
function showWindow()
{
    var win = new Window("palette","Batch Liste de travail",undefined,{resizeable:false});
    //var rootPanel = win.add("panel");
    
    // top group
    var groupUp = win.add("group");
    // GREP  Panel
    var panelGreps = groupUp.add("panel");    
    var grepLabel = panelGreps.add("statictext");
    grepLabel.text = "liste GREP";
	grepLabel.helpTip = "Clic-droit pour ajouter";
    grepList = panelGreps.add("listbox");
    grepList.size = {width:200,height:500};
    if(grepFiles!=null){
        for(var v1=0;v1<grepFiles.length;v1++){
            grepList.add("item",grepFiles[v1].displayName);
        }
        grepList.addEventListener("click", function(k){doubleClickGrepItem(k);});
    }
    // Scripts Panel
    var panelScripts = groupUp.add("panel");
    var scriptsLabel = panelScripts.add("statictext");
    scriptsLabel.text = "Liste Script";
	scriptsLabel.helpTip = "Clic-droit pour ajouter";
    scriptsList = panelScripts.add("listbox");
    scriptsList.size = {width:200,height:500};
    if(scriptFiles!=null){
        for(var v2=0;v2<scriptFiles.length;v2++){
            scriptsList.add("item",scriptFiles[v2].displayName);
        }
        scriptsList.addEventListener("click", function(k){doubleClickScriptItem(k);});
    }
    var panelJobs = groupUp.add("panel");
    var jobsLabel = panelJobs.add("statictext");
    jobsLabel.text = "Vos files d'attente";
	jobsLabel.helpTip = "SUPPR pour supprimer de la liste";
    jobsList = panelJobs.add("listbox",undefined,undefined,{multiselect:false});
    jobsList.size = {width:200,height:400};
    jobsList.addEventListener ("keydown", function(kd){pressed(kd);});    

    // jobs button group 1
    var buttonGroup1 = panelJobs.add("group");
    var moveUpButton = buttonGroup1.add("button",undefined,"Monter");// job move up
    moveUpButton.onClick = function(){moveItemUp(jobsList);};//Déplacer des événements de bouton
    var moveDownButton = buttonGroup1.add("button",undefined,"Descendre");//job move down
    moveDownButton.onClick = function(){moveItemDown(jobsList);};//en bas événement Bouton
    // jobs button group 2
    var buttonGroup2 = panelJobs.add("group");
    var exportButton = buttonGroup2.add("button",undefined,"Exporter la liste");
    exportButton.onClick = function(){exportJobsList();};//Exporter Annonce
    var importButton = buttonGroup2.add("button",undefined,"Importer une liste");
    importButton.onClick = function(){importJobsList();};//Exporter Annonce
    
    // jobs button group 3
    var buttonGroup3 = panelJobs.add("group");
    var runButton = buttonGroup3.add("button",undefined,"Exécuter la liste");
    runButton.onClick = function(){runJobsList();};//Exporter Annonce
    
    // DOWN GROUP
    //fenêtre de message
    var groupDown = win.add("group");
    msgText = groupDown.add("edittext",[0,0,700,100],"",{multiline:true});
    
    win.addEventListener ("keydown", function(kd){pressed(kd);});//Fenêtres Appuyez sur DEL pour supprimer JobsList directement représenté dans le projet
    win.show();
}

//==================================== Function Level 1 ============================================

//get scripts location
function script_folder()
{
	try {
        return File (app.activeScript).path+"/";
    }catch(e) {
        return File (e.fileName).path+"/";
    }
}

//get grep file collection
function getGrepFiles()
{
    var queryFolder = app.scriptPreferences.scriptsFolder.parent.parent + "/Find-Change Queries/Grep/";
    //alert(queryFolder);
    var grepFolder = new Folder(queryFolder);
    var grepFilesCollect = new Array();//Archives de la collection de GREP
    //Extension de fichier XML doit être
    if(grepFolder.exists){
        var filesCollect = grepFolder.getFiles();//Toutes les archives non filtrée
        var iCount = 0;
        for(var v1=0;v1<filesCollect.length;v1++){
            var ext = getExtFileName(filesCollect[v1].name);
            if(ext=="xml"){
                grepFilesCollect[iCount] = filesCollect[v1];
                iCount += 1;
            }
        }
    }
    if(grepFilesCollect.length>0){
        return grepFilesCollect;
    }
    return null;
}

//get script file collection
function getScriptFiles()
{
    var queryFolder = app.scriptPreferences.scriptsFolder;
    var scriptFolder = new Folder(queryFolder);
    var scriptFilesCollect = new Array();//Collection de fichiers de script
    //File extension doit être jsx ou jsxbin
    if(scriptFolder.exists){
        var filesCollect = scriptFolder.getFiles();//Toutes les archives non filtrée
        var iCount = 0;
        for(var v1=0;v1<filesCollect.length;v1++){
            var ext = getExtFileName(filesCollect[v1].name);
            if(ext=="jsx" || ext=="jsxbin"){
                scriptFilesCollect[iCount] = filesCollect[v1];
                iCount += 1;
            }
        }
    }
    if(scriptFilesCollect.length>0){
        return scriptFilesCollect;
    }
    return null;
}

function doubleClickGrepItem(p)
{    
    //var nowListBox = p.target;
    //alert(nowListBox.selection);
    jobsList.add("item",grepList.selection);
}

function doubleClickScriptItem(p)
{    
    //var nowListBox = p.target;
    //alert(nowListBox.selection);
    jobsList.add("item",scriptsList.selection);
}

// delete jobs
function pressed(k)
{
  switch(k.keyName)
    {
        case "Delete":
            if(jobsList.selection==null){
                jobsList.selection = [0];
                //alert("selection null");
            }else{
                try{
                    var mySelection = jobsList.selection;//Les projets sélectionnés
                    jobsList.remove(mySelection);//Supprimer des éléments
                }catch(e){
                    alert(e.message);
                }
            }
        break;
        default: //nothing
            //alert("other");
    }
}

//projet déménagement
function moveItemUp(nowListBox)
{
    if(nowListBox.selection!=null && nowListBox.selection.index>0)
    {
        var n = nowListBox.selection.index;
        swap(nowListBox.items[n],nowListBox.items[n-1]);
        nowListBox.selection = n-1;    
    }
}

//en bas projet
function moveItemDown(nowListBox)
{
    if(nowListBox.selection!=null && nowListBox.selection.index<nowListBox.items.length-1)
    {
        var n = nowListBox.selection.index;
        swap(nowListBox.items[n],nowListBox.items[n+1]);
        nowListBox.selection = n+1;    
    }
}

//export jobs list
function exportJobsList()
{
    var saveFile = File.saveDialog (" Enregistrer le fichier", "*.job");
    saveFile.encoding = "UTF-8";
    saveFile.open("w");
    for(var i=0;i<jobsList.items.length;i++){
        saveFile.writeln(jobsList.items[i].text+"\r");
    }
    saveFile.close();
}

//import jobs list
function importJobsList()
{
    var openFile = new File(File.openDialog("liste d'importation", "*.job", "false"));
    var filesArray = new Array();
    var iCount = 0;
    if(openFile.open("r")==true){
        while(openFile.eof!=true){
            var line = openFile.readln();
            filesArray[iCount] = line;
            iCount = iCount+1;
        }
    }
    if(filesArray.length>0 && filesArray!=null ){
        jobsList.removeAll();
        for(var v1=0;v1<filesArray.length;v1++){
            jobsList.add ("item", filesArray[v1]);        
        }
    }
    openFile.close();
}

function swap(x,y)
{
    var temp = x.text;
    x.text = y.text;
    y.text = temp;
}

function runJobsList()
{
    // Pas de liste vide
    if(jobsList.items.length<=0){
        return null;
    }
    var jobs = jobsList.items;//Recueillir tout le travail
    //GREP juge ou un script
    for(var i=0;i<jobs.length;i++){
        var ary = jobs[i].text.split (".");
        var extFileName = ary[ary.length-1];//Le dernier est l'extension de fichier
        switch(extFileName){
            case "xml":
                msgText.text += "[GREP]" + jobs[i].text + "\r\n";
                runGrep(jobs[i].text);// RUN GREP
            break;
            case "jsx":
                msgText.text += "[jsx]" + jobs[i].text + "\r\n";
                runScript(jobs[i].text);//RUN SCRIPT
            break;
            case "jsxbin":
                msgText.text += "[jsxbin]" + jobs[i].text + "\r\n";
                runScript(jobs[i].text);//RUN BIN SCRIPT
            break;
            default:
                msgText.text += "===[undefined]: " + jobs[i].text + " n'est pas un fichier valide ! ===\r\n";
        }
    }
}

//===================================== Function Level 2 ==============================================

//EXECUTE GREP
function runGrep(fileName)
{
    var queryFolder = app.scriptPreferences.scriptsFolder.parent.parent + "/Find-Change Queries/Grep/";
    var grepFile = new File(queryFolder + fileName);
    if(grepFile.exists==false){
        //File does not exist
        msgText.text += fileName + " n'existe pas ! \r\n";
        return null;
    }
    var inDoc = app.activeDocument;
    var subEndNum = fileName.lastIndexOf (".xml", fileName.length);
    var grepName = fileName.substring (0,subEndNum);
    app.loadFindChangeQuery(grepName,SearchModes.grepSearch);
    inDoc.changeGrep();
}

//EXECUTE SCRIPT
function runScript(fileName){
    var scriptFile = new File(app.scriptPreferences.scriptsFolder+"/"+fileName);
    if(scriptFile.exists==false){
        //File does not exist
        msgText.text += fileName + " n'existe pas ! \r\n";
        return null;
    }
    var result = app.doScript(scriptFile);
    msgText.text += result + "\r\n";
}

//EXECUTE BIN SCRIPT
//~ function runBinScript(fileName){
//~     var scriptFile = new File(app.scriptPreferences.scriptsFolder+"/"+fileName);
//~     if(scriptFile.exists==false){
//~         //File does not exist
//~         msgText.text += fileName + " n'existe pas ! \r\n";
//~         return null;
//~     }
//~     var result = app.doScript(scriptFile);
//~     msgText.text += result + "\r\n";
//~ }

//progress bar
function progressBarMake(parentWin,stop){
    var pbar = parentWin.add("progressbar",undefined,1,stop);
    pbar.preferredSize = [300,20];
    parentWin.show();
    return pbar;
}

//Obtenir l'extension de fichier
function getExtFileName(fileName)
{
    var splitAry = fileName.split(".");
    var extName = splitAry[splitAry.length-1];//La dernière extension de fichier doit être
    return extName;
}
