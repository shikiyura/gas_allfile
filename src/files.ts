type FileName = string;
type OwnerAddress = string;
type FolderName = string;
type Url = string;
type FileInfo = [FileName, OwnerAddress, FolderName, Url];


let limitFlag = false;

export let sheetWrite = (sheetName: string, data: any[][]) => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if(sheet === null) return;

  const lastRow = sheet.getLastRow();
  console.log(lastRow, data);

  if(data.length === 0 || data[0].length === 0) {
    console.log("データなし？");
    return;
  }

  const range = sheet.getRange(lastRow + 1, 1, data.length, data[0].length);
  range.setValues(data);
}

export let allFolders = () => {
  const startDate = new Date();
  const functionName = "allFolders"
  let subFolders: GoogleAppsScript.Drive.FolderIterator[] = [];
  const files: FileInfo[] = [];

  let resumeData = getResume(functionName);
  if (resumeData === null) {
    const url = "https://drive.google.com/drive/u/0/folders/0B7EYQ1nXcfKEb3dwVnlBUk9SVFE";
    const targetDir = DriveApp.getFolderById(url.split("folders/")[1]);
    subFolders = [targetDir.getFolders()];
  } else {
    const resume = resumeData.split("???");
    console.log(resume);
    const folderIterator = DriveApp.continueFolderIterator(resume[0]);
    subFolders.push(folderIterator);

    if (resume[1].length !== 0) {
      subFolders.push(...JSON.parse(resume[1]))
    }
  }

  for(let i = 0; i < subFolders.length; i++) {
    const subFolder = subFolders[i];

    while (subFolder.hasNext()) {
      const folder = subFolder.next();
      const subFiles = allFiles(folder);

      files.push(...Array.from(subFiles));
      subFolders.push(folder.getFolders());

      if(limitFlag || checkLimit(startDate, functionName)) {
        setResume(functionName, subFolder.getContinuationToken(), JSON.stringify(subFolders));
        break;
      };
    }

    if(limitFlag || checkLimit(startDate, functionName)) {
      setResume(functionName, subFolder.getContinuationToken());
      break;
    };
  }

  sheetWrite("ファイル一覧", files);
}

export let allFiles = (folder: GoogleAppsScript.Drive.Folder | null): FileInfo[] => {
  const startDate = new Date();
  let folderName = null, fileIterator;
  const files: FileInfo[] = [];
  if (folder === null) {
    const scriptProperty = PropertiesService.getScriptProperties();
    const token = scriptProperty.getProperty("allFiles")!;
    fileIterator = DriveApp.continueFileIterator(token);
  } else {
    folderName = folder.getName();
    fileIterator = folder.getFiles();
  }


  while (fileIterator.hasNext()) {
    const file = fileIterator.next();

    files.push([
      file.getName(),
      file.getOwner().getEmail(),
      folderName || file.getParents().next().getName(),
      file.getUrl()
    ]);

    if(checkLimit(startDate, "allFiles")) {
      setResume("allFiles", fileIterator.getContinuationToken());
      break;
    }
  }

  return files;
}

const checkLimit = (startDate: Date, calledFunction: string): boolean => {
  const nowDate = new Date();
  const proc_secs = (nowDate.getTime() - startDate.getTime()) / 1000;

  limitFlag = (proc_secs >= 300)

  if(limitFlag) {
    setTrigger(calledFunction);
  }

  return limitFlag;
}

const setResume = (target: string, iteratorToken: string, appendix: string = "") => {
  const scriptProperty = PropertiesService.getScriptProperties();
  scriptProperty.setProperty(target, `${iteratorToken}???${appendix}`);
}

const getResume = (target: string) => {
  const scriptProperty = PropertiesService.getScriptProperties();
  const property = scriptProperty.getProperty(target);
  scriptProperty.deleteProperty(target);

  return property;
}

const setTrigger = (functionName: string) => {
  let triggers = ScriptApp.getProjectTriggers();
  for(let trigger of triggers) {
    if(trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  ScriptApp.newTrigger(functionName).timeBased().after(1000 * 60).create();
}
