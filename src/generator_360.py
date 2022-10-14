from lib2to3.pgen2 import token
from IPython.display import clear_output
import ipywidgets as widgets
import pandas as pd
import time
import re
import numpy as np
from IPython.display import Markdown
import webbrowser


class Generator:
    def __init__(self, dcms, os, auth, requestList, serviceSheet):
        self.dcms = dcms
        self.os = os
        self.requestList = requestList
        self.serviceSheet = serviceSheet
        self.auth = auth
        self.clientName = None
        self.projectName = None
        self.num_level = None
        self.marker_count = 1
        self.level_count = 0
        self.columnformarker = 0
        self.level_flag = False
        self.columnforlevels = 7
        self.rowsforlevels = 3
        self.toplevel = ""
        self.rowForTabsName = None
        self.toplevelID = None
        self.sameLeveID = None
        self.linkFigmaData = None
        self.linkFigmaSheet = None
        self.linkFigmaSheetName = "linkFigma"
        self.updatesheetname = "updatesheet"
        self.projectSheetId = "10mWRMQn2KBY7HxCGNJtk1--U5FcLmD99Nv9-hbJC8Aw"
        self.projectData = auth.open_by_key(self.projectSheetId)
        self.teamSheet = None
        self.teamSheetDataDF = None
        self.omitted_levels = None
        self.updatesheetid = None
        self.taskbreakdownid = None
        self.taskbreakdownsheet_name = None
        self.exclude = None
        self.gettingInformation(self.auth)

    def connectDCMS(self):

        DISPERSE_TOKEN = self.os.getenv("DISPERSE_TOKEN_US")
        self.cloud = self.dcms(
            username=self.os.getenv("DISPERSE_USERNAME"),
            password=self.os.getenv("DISPERSE_PASSWORD"),
            token=DISPERSE_TOKEN,
        )
        clear_output(wait=True)
        display(Markdown("<center><h4>!!!!Connecting you to DCMS!!!!!<h4></center> "))
        self.cloud.connect(
            username=self.os.getenv("DISPERSE_USERNAME"),
            password=self.os.getenv("DISPERSE_PASSWORD"),
            customer="internal-" + self.clientName,
            project=self.projectName,
        )

    def setConstants(self):
        refCol = self.columnforlevels
        refRow = self.rowsforlevels
        self.marker_count = 1
        return refRow, refCol

    def lastRun(self):
        self.connectDCMS()
        self.setConstants()
        self.gettingSheets(self.auth)
        self.gettingLevels(time)
        self.populatingLinkFigma()
        self.UpdatingSheet()

    def handle_submit(self, change):
        if len(self.omitted_levels.value) > 0:
            self.omitted_levels = self.omitted_levels.value
        self.main_output.clear_output()
        clear_output(wait=True)
        self.lastRun()

    def get_centered_layout(self, display_width="50%"):
        return widgets.Layout(
            display="flex", flex_flow="row", align_items="center", width=display_width
        )

    def needdivision(self, change):
        if change.new == "Yes":
            display(Markdown("<h5 style='margin-left:4%;'>Level Names?</h5>"))
            self.omitted_levels = widgets.Text(
                value="", description=" ", layout=self.get_centered_layout()
            )
            self.omitted_levels.on_submit(self.handle_submit)
            with self.main_output:
                display(self.omitted_levels)
        else:
            self.main_output.clear_output()
            clear_output(wait=True)
            self.lastRun()

    def selectProject(self, change):
        display("Project Selected is: " + change.new)
        self.rowForTabsName = int(
            self.teamSheetDataDF[self.teamSheetDataDF["PROJECT"] == change.new].index[0]
        )
        SelectedProject = self.teamSheetDataDF[
            self.teamSheetDataDF["PROJECT"] == change.new
        ].reset_index(drop=True)
        tabColumns = SelectedProject[
            [col for col in SelectedProject.columns if "TAB" in col]
        ]
        self.upsTabNames = [val for val in tabColumns.values.tolist()[0] if val != ""]
        self.clientName = SelectedProject["CUSTOMER NAME"][0]
        self.projectName = SelectedProject["PROJECT NAME"][0]
        self.num_level = int(SelectedProject["DEPTH"][0])
        self.updatesheetid = SelectedProject["UPDATE SHEET ID"][0]
        self.taskbreakdownid = SelectedProject["TASKBREAKDOWN SHEET ID"][0]
        self.taskbreakdownsheet_name = SelectedProject["TASKBREAKDOWN SHEET NAME"][0]
        self.toplevelID = SelectedProject["MAIN LEVEL ID"][0]
        self.tableDivisionLevels = (
            [
                x.title().strip()
                for x in SelectedProject["Table Division Levels"][0].split(",")
                if x != ""
            ]
            if SelectedProject["Table Division Levels"][0] != ""
            else None
        )
        display(Markdown("<h5 style='margin-left:4%;'>Omit levels?</h5>"))
        omit_levels = widgets.Dropdown(
            options=["Yes", "No"],
            description=" ",
            value=None,
            layout=self.get_centered_layout(),
        )
        omit_levels.observe(self.needdivision, names="value")
        with self.main_output:
            display(omit_levels)

    def selectTeam(self, change):
        clear_output(wait=True)
        display("Team Selected is: " + change.new)
        self.teamSheet = self.projectData.worksheet_by_title(change.new)
        self.teamSheetDataDF = self.teamSheet.get_as_df(
            include_tailing_empty=True, include_tailing_empty_rows=False
        )
        display(Markdown("<h5 style='margin-left:4%;'>Select Project: </h5>"))
        projectSelected = widgets.Dropdown(
            options=self.teamSheetDataDF["PROJECT"].tolist(),
            value=None,
            description=" ",
            layout=self.get_centered_layout(),
        )
        projectSelected.observe(self.selectProject, names="value")
        with self.main_output:
            display(projectSelected)

    def gettingInformation(self, authorizing):
        display(Markdown("<center><h3>Updatesheet Generator</h3></center>"))
        allTeams = []
        [allTeams.append(name.title) for name in self.projectData.worksheets()]
        display(Markdown("<h5 style='margin-left:4%;'>Select Your Team: </h5>"))
        teamSelected = widgets.Dropdown(
            options=allTeams,
            value=None,
            description=" ",
            layout=self.get_centered_layout(),
        )
        teamSelected.observe(self.selectTeam, names="value")
        self.main_output = widgets.Output()
        display(self.main_output)
        with self.main_output:
            display(teamSelected)

    def gettingSheets(self, authorizing):
        clear_output(wait=True)
        display(
            Markdown("<center><h4>Getting information form the Sheets</center></h4>")
        )
        self.taskBreakDownSpreadSheet = authorizing.open_by_key(self.taskbreakdownid)
        self.TBworkingSheet = self.taskBreakDownSpreadSheet.worksheet_by_title(
            self.taskbreakdownsheet_name
        )
        self.updateSpreadSheet = authorizing.open_by_key(self.updatesheetid)
        try:
            self.linkFigmaSheet = self.updateSpreadSheet.worksheet_by_title(
                self.linkFigmaSheetName
            )
        except:
            self.linkFigmaSheet = self.updateSpreadSheet.add_worksheet(
                "linkFigma", rows=500, cols=7
            )
        self.TBworkingSheetDF = pd.DataFrame(
            self.TBworkingSheet.get_all_values(
                returnas="matrix",
                majdim="ROWS",
                include_tailing_empty=False,
                include_tailing_empty_rows=False,
            )
        )

        self.findingandsettingData()

    def findingandsettingData(self):
        clear_output(wait=True)
        display(Markdown("<center><h4>Got the Information.</center></h4>"))
        display(
            Markdown(
                "<center><h4>Now Finding Some Keywords From the Sheet</center></h4>"
            )
        )
        global dict_notes
        findingTasks = self.TBworkingSheet.find("TASK GROUP")
        findingComponents = self.TBworkingSheet.find("Components - by Disperse")
        # ! ------------------------------- for Adding Notes ----------------------------- !
        findingUID = self.TBworkingSheet.find("UID")
        findingComments = self.TBworkingSheet.find("Clarifications - by Disperse")
        datatoPaste = self.TBworkingSheetDF.loc[
            findingTasks[0].row + 1 :,
            findingTasks[0].col - 2 : findingComponents[0].col,
        ]
        # ! ------------------------------------------------------------------------------- !
        # ! ------------------------------------- for Adding Notes ------------------------------- !
        notesDataFrame = self.TBworkingSheetDF[
            [findingUID[-1].col - 1, findingComments[0].col - 1]
        ].dropna()
        notesDataFrame.columns = ["Uid", "Comments"]
        dict_notes = notesDataFrame.set_index("Uid").T.to_dict("list")
        self.filteringData(datatoPaste)
        # ! --------------------------------------------------------------------------------------- !

    def filteringData(self, datarecevied):
        global datatoPaste
        clear_output(wait=True)
        display(Markdown("<center><h4>Dropping Extra Information</center></h4>"))
        datatoPaste = datarecevied.drop(
            [datarecevied.columns.stop - 2, datarecevied.columns.start + 2], axis=1
        )
        datatoPaste.columns = ["Task Uid", "task_name", "Components Id", "task"]
        datatoPaste = datatoPaste[
            ["Task Uid", "Components Id", "task_name", "task"]
        ]  # rearranging the columns
        datatoPaste = datatoPaste[datatoPaste["task"].str.lower() != "non-trackable"]
        self.linkFigmaData = datatoPaste[["Task Uid", "task_name"]]
        datatoPasteColumn = list(datatoPaste.columns)  # converting into list
        datatoPaste = datatoPaste.values.tolist()
        datatoPasteColumn.append("dpnd")
        datatoPaste.insert(0, datatoPasteColumn)

    def transversingLevels(self, data, levelsData, parents):
        for levels in data:
            if levels["name"] != "Objects":
                parents.append(levels["name"] if "name" in levels else "#")
                if len(levels["children"]) > 0:
                    levelsData = self.transversingLevels(
                        levels["children"], levelsData, parents
                    )
                levelsData.append(
                    [
                        levels["name"] if "name" in levels else "",
                        levels["_id"],
                        levels["excluded"] if "excluded" in levels else "",
                        levels["template_type"] if "tempplate_type" in levels else "",
                        parents.copy(),
                    ]
                )
                parents.pop()
        return levelsData

    def gettingLevels(self, time):
        allLevels = self.cloud.get_levels_tree()
        df_tree = pd.DataFrame(
            self.transversingLevels(allLevels["children"], [], ["#"]),
            columns=["name", "id", "excluded", "template_type", "ancestors"],
        )
        df_tree = pd.concat(
            [
                df_tree,
                pd.DataFrame(df_tree["ancestors"].values.tolist()).add_prefix(
                    "parent_"
                ),
            ],
            axis=1,
        )
        newdftree = df_tree[[col for col in df_tree if col.startswith("parent")]]
        newdftree = self.filteringLevels(newdftree, self.num_level + 1)
        if self.omitted_levels is not None:
            display(Markdown("<center><h4>Omitting levels</center></h4>"))
            newdftree = newdftree[
                ~(
                    newdftree.isin(
                        map(lambda x: x.title().strip(), self.omitted_levels.split(","))
                    ).any(axis=1)
                )
            ]
        newdftree = (
            newdftree.drop(["parent_0", "parent_1"], axis=1)
            .iloc[:, : self.num_level]
            .fillna(value="")
        )
        newdftree["combined_level"] = newdftree.agg(
            lambda row: "<>".join([x for x in row if x != ""]), axis=1
        )
        newdftree = newdftree.drop_duplicates(subset=["combined_level"]).drop(
            ["combined_level"], axis=1
        )

        display(newdftree)
        fgdfgdf
        for tab in self.upsTabNames:
            if tab.lower() == "updatesheet" or tab.lower() == "update sheet":
                level_final = newdftree.replace(to_replace="Objects", value="~")
            else:
                level_final = newdftree[newdftree.isin([tab]).any(axis=1)].replace(
                    to_replace="Objects", value="~"
                )
            LEVELSTOPASTE = {}
            for lst in level_final[level_final.columns].values:
                leaf = LEVELSTOPASTE
                for path in lst[:-2]:
                    if path != "":
                        leaf = leaf.setdefault(path, {})
                if lst[-2] != "" and lst[-1] != "":
                    leaf.setdefault(lst[-2], list()).append(lst[-1])
            try:
                newSheet = self.updateSpreadSheet.add_worksheet(
                    tab, rows=10000, cols=104
                )
            except:
                self.updateSpreadSheet.del_worksheet(
                    self.updateSpreadSheet.worksheet_by_title(tab)
                )
                newSheet = self.updateSpreadSheet.add_worksheet(
                    tab, rows=10000, cols=104
                )
            rowforlevels, columnforlevels = self.setConstants()
            self.retransversing(
                LEVELSTOPASTE,
                rowforlevels,
                columnforlevels,
                newSheet.id,
                self.level_flag,
                {},
                {},
                {},
                {},
                self.level_count,
                0,
            )

    def filteringLevels(self, referenceLeveltreeData, totalParent):
        if self.toplevelID[0] == "-" or self.toplevelID[0] == "":
            return referenceLeveltreeData[
                referenceLeveltreeData["parent_" + str(totalParent)] != ""
            ].reset_index(drop=True)
        else:
            referenceLeveltreeData = referenceLeveltreeData[
                referenceLeveltreeData.isin(
                    [self.cloud.get_object({"id": self.toplevelID})["name"]]
                ).any(axis=1)
            ]
            return referenceLeveltreeData.reset_index(drop=True)

    def gettingSublevel(self, referenceSublevel, referenceSubLeveltreeData):
        sublevelDataframe = pd.DataFrame()
        for sublevel in referenceSublevel:
            sublevelDataframe = pd.concat(
                [
                    sublevelDataframe,
                    referenceSubLeveltreeData[
                        referenceSubLeveltreeData.isin(
                            [self.cloud.get_object({"id": sublevel})["name"]]
                        ).any(axis=1)
                    ],
                ]
            )
        return sublevelDataframe

    def makingRequest(self, valueToAdd, sheetID, rowToadd, columnToAdd):
        self.requestList.append(
            {
                "updateCells": {
                    "rows": [
                        {"values": [{"userEnteredValue": {"stringValue": valueToAdd}}]}
                    ],
                    "fields": "userEnteredValue.stringValue",
                    "start": {
                        "sheetId": sheetID,
                        "rowIndex": rowToadd,
                        "columnIndex": columnToAdd,
                    },
                }
            }
        )

    def makingNotesRequest(self, valueToAdd, sheetID, rowToadd, columnToAdd):
        self.requestList.append(
            {
                "updateCells": {
                    "rows": [{"values": [{"note": valueToAdd}]}],
                    "fields": "note",
                    "start": {
                        "sheetId": sheetID,
                        "rowIndex": rowToadd,
                        "columnIndex": columnToAdd,
                    },
                }
            }
        )

    def addTaskRows(self, datatoPaste, sheetID, rowforlevel, columnformarker):
        for data in datatoPaste:
            COLFORTASKS = 1
            for entry in data:
                if entry != "" and entry in dict_notes:
                    self.makingNotesRequest(
                        dict_notes[entry][0], sheetID, rowforlevel, 4
                    )
                self.makingRequest(
                    re.sub(" +", " ", entry).strip(), sheetID, rowforlevel, COLFORTASKS
                )
                if entry == "dpnd":
                    self.makingRequest(
                        "^" + str(self.marker_count),
                        sheetID,
                        rowforlevel,
                        COLFORTASKS + 1,
                    )
                COLFORTASKS = COLFORTASKS + 1
            rowforlevel = rowforlevel + 1
        self.makingRequest(
            "^" + str(self.marker_count), sheetID, rowforlevel, columnformarker
        )
        self.marker_count += 1
        return rowforlevel

    def retransversing(
        self,
        LevelsData,
        rowforlevels,
        columnforlevels,
        sheetId,
        level_flag,
        newSheetsNames,
        parentDict,
        levelDict,
        markerDict,
        levelCount,
        rowForTasks,
    ):
        for Level in LevelsData:
            if Level != "~":
                parentDict[Level] = [rowforlevels, columnforlevels]
                levelDict["level_" + str(levelCount)] = [rowforlevels, 6]
                rowforlevels += 1
                levelCount += 1
                if isinstance(LevelsData[Level], dict):
                    rowforlevels, columnforlevels, rowForTasks = self.retransversing(
                        LevelsData[Level],
                        rowforlevels,
                        columnforlevels,
                        sheetId,
                        level_flag,
                        newSheetsNames,
                        parentDict,
                        levelDict,
                        markerDict,
                        levelCount,
                        rowForTasks,
                    )
                elif isinstance(LevelsData[Level], list):
                    levelDict["level_" + str(levelCount)] = [rowforlevels, 6]
                    if rowForTasks < rowforlevels:
                        rowForTasks = rowforlevels
                    self.pastingDictionaries(parentDict, sheetId)
                    self.pastingDictionaries(levelDict, sheetId)
                    self.makingRequest(
                        "^" + str(self.marker_count),
                        sheetId,
                        levelDict["level_0"][0] - 1,
                        3,
                    )
                    for lastLevel in LevelsData[Level]:
                        self.makingRequest(
                            lastLevel, sheetId, rowforlevels, columnforlevels
                        )
                        columnforlevels += 1
                    del (
                        parentDict[Level],
                        levelDict["level_" + str(levelCount)],
                        levelDict["level_" + str(levelCount - 1)],
                    )
                    if (Level.lower().find("Floor".lower()) < 0) & (
                        Level.lower().find("Basement".lower()) < 0
                    ):
                        levelCount = levelCount - 1
                    rowforlevels -= 1
                if self.tableDivisionLevels is None:
                    if (
                        (Level.lower().find("Floor".lower()) >= 0)
                        | (Level.lower().find("Basement".lower()) >= 0)
                        | (Level.lower().find("Mezzanine".lower()) >= 0)
                        | (Level.lower().find("Roof".lower()) >= 0)
                        | (Level.lower().find("Plantroom".lower()) >= 0)
                    ):
                        if (Level in parentDict) & (
                            ("level_" + str(levelCount - 1)) in levelDict
                        ):
                            parentDict.pop(Level, None), levelDict.pop(
                                "level_" + str(levelCount - 1), None
                            )
                        rowforlevels = self.addTaskRows(
                            datatoPaste, sheetId, rowForTasks + 3, columnforlevels
                        )
                        self.resetDictForFloor(rowforlevels + 5, 7, parentDict)
                        rowforlevels = self.resetDictForFloor(
                            rowforlevels + 5, 6, levelDict
                        )
                        levelCount -= 1
                        columnforlevels = 7
                else:
                    if Level in self.tableDivisionLevels:
                        if (Level in parentDict) & (
                            ("level_" + str(levelCount - 1)) in levelDict
                        ):
                            parentDict.pop(Level, None), levelDict.pop(
                                "level_" + str(levelCount - 1), None
                            )
                        rowforlevels = self.addTaskRows(
                            datatoPaste, sheetId, rowForTasks + 3, columnforlevels
                        )
                        self.resetDictForFloor(rowforlevels + 5, 7, parentDict)
                        rowforlevels = self.resetDictForFloor(
                            rowforlevels + 5, 6, levelDict
                        )
                        levelCount -= 1
                        columnforlevels = 7
                if (Level in parentDict) & (
                    ("level_" + str(levelCount - 1)) in levelDict
                ):
                    parentDict.pop(Level, None), levelDict.pop(
                        "level_" + str(levelCount - 1), None
                    )
                    levelCount = levelCount - 1
                    rowforlevels -= 1

        return rowforlevels, columnforlevels, rowForTasks

    def pastingDictionaries(self, refereneceDict, sheetID):
        for item in refereneceDict:
            if item != "sheetID":
                self.makingRequest(
                    item, sheetID, refereneceDict[item][0], refereneceDict[item][1]
                )

    def resetDictForFloor(self, referenceRow, referenceCol, referenceDict):
        for keys in referenceDict:
            if keys != "sheetID":
                referenceDict[keys][0] = referenceRow
                referenceDict[keys][1] = referenceCol
                referenceRow += 1
        return referenceRow

    def UpdatingSheet(self):
        clear_output(wait=True)
        display(
            Markdown("<center><h4>Finally Pasting Everything Everywhere</h4></center>")
        )
        requestBody = {"requests": self.requestList}
        result = (
            self.serviceSheet.spreadsheets()
            .batchUpdate(spreadsheetId=self.updatesheetid, body=requestBody)
            .execute()
        )
        clear_output(wait=True)
        display(
            Markdown(
                "<h3><center>Hurrah!!!! Your update Sheet has been Created. Redirecting....</center></h3>"
            )
        )
        webbrowser.open(
            "https://" + "docs.google.com/spreadsheets/d/" + str(self.updatesheetid)
        )

    def populatingLinkFigma(self):
        clear_output(wait=True)
        display(Markdown("<center><h4>Populating linkFigma Sheet </center></h4>"))
        self.linkFigmaData = self.linkFigmaData[
            self.linkFigmaData[["Task Uid", "task_name"]].ne("").all(axis=1)
        ].values.tolist()
        linkfigmaRow = 1
        for data in self.linkFigmaData:
            linkfigmaColumn = 0
            for task in data:
                self.makingRequest(
                    task, self.linkFigmaSheet.id, linkfigmaRow, linkfigmaColumn
                )
                linkfigmaColumn += 1
            linkfigmaRow += 1
