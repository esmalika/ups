from ast import Continue
from lib2to3.pgen2 import token
from math import floor
from sys import displayhook
from IPython.display import clear_output, display
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
        self.region = None
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
        self.dcmsius_stack_parent = []
        self.rowForTabsName = None
        self.toplevelID = None
        self.sameLeveID = None
        self.linkFigmaData = None
        self.linkFigmaSheet = None
        self.linkFigmaSheetName = "linkFigma"
        self.updatesheetname = "updatesheet"
        self.projectSheetId = "10mWRMQn2KBY7HxCGNJtk1--U5FcLmD99Nv9-hbJC8Aw"
        self.projectData = self.auth.open_by_key(self.projectSheetId)
        self.teamSheet = None
        self.teamSheetDataDF = None
        self.omitted_levels = None
        self.updatesheetid = None
        self.taskbreakdownid = None
        self.taskbreakdownsheet_name = None
        self.exclude = None
        self.selectRegion()
        # self.gettingInformation(self.auth)

    def connectDCMS(self):
        kwargs = {
            "username": self.os.getenv("DISPERSE_USERNAME"),
            "password": self.os.getenv("DISPERSE_PASSWORD"),
            "token": self.os.getenv("DISPERSE_TOKEN_" + self.region.upper())
            if self.region is not None
            else None,
            "projects_config_path": "G:\Shared drives\GENERATORS\special_files\migrated_projects.json",
        }

        self.cloud = self.dcms(**kwargs)
        clear_output(wait=True)
        display(Markdown("<center><h4>!!!!Connecting you to DCMS!!!!!<h4></center> "))
        try:
            self.cloud.connect(
                customer=self.clientName,
                project=self.projectName,
                region=self.region,
                scope="internal",
                environment="production",
            )
        except:
            display(
                Markdown(
                    "<center><h2>Couldn't Connect to DC, either project is Shell or region selected is wrong</center></h2>"
                )
            )
            exit()
        display("Everything connected!")

    def get_migrated_json_path(self):
        list_of_levels_up = []
        for i in range(1, 11):
            temp_list_of_levels_up = i * [".."]
            if self.os.path.isfile(
                self.os.path.join(
                    *temp_list_of_levels_up,
                    "GENERATORS",
                    "special_files",
                    "migrated_projects.json"
                )
            ):
                list_of_levels_up = temp_list_of_levels_up
                break
        return self.os.path.join(
            *list_of_levels_up, "GENERATORS", "special_files", "migrated_projects.json"
        )

    def setConstants(self):
        refCol = self.columnforlevels
        refRow = self.rowsforlevels
        self.marker_count = 1
        return refRow, refCol

    def get_centered_layout(self, display_width="100%"):
        return widgets.Layout(
            display="flex",
            flex_flow="column",
            align_items="center",
            width=display_width,
        )

    def selectRegion(self):
        display(Markdown("<center><h3>Updatesheet Generator</h3></center>"))
        regions = ["uk", "us", "de", "au"]
        regionSelected = widgets.Dropdown(
            options=regions,
            value=None,
            description=" ",
            layout=self.get_centered_layout(),
        )
        regionSelected.observe(self.gettingInformation, names="value")
        self.main_output = widgets.Output()
        region_dropdown_box = widgets.HBox(
            children=[regionSelected], layout=self.get_centered_layout()
        )
        display(self.main_output)
        with self.main_output:
            display(region_dropdown_box)

    def gettingInformation(self, change):
        clear_output(wait=True)
        self.region = change.new
        display(Markdown("<center><h5>Region: {}<h5></center> ".format(self.region)))
        allTeams = []
        [allTeams.append(name.title) for name in self.projectData.worksheets()]
        # display(Markdown("<h5 style='margin-left:4%;'>Select Your Team: </h5>"))
        teamSelected = widgets.Dropdown(
            options=allTeams,
            value=None,
            description=" ",
            layout=self.get_centered_layout(),
        )
        teamSelected.observe(self.selectTeam, names="value")
        with self.main_output:
            display(teamSelected)

    def selectTeam(self, change):
        clear_output(wait=True)
        display(Markdown("<center><h5>Region: {}<h5></center> ".format(self.region)))
        display(Markdown("<center><h5>Team: {}<h5></center> ".format(change.new)))
        self.teamSheet = self.projectData.worksheet_by_title(change.new)
        self.teamSheetDataDF = self.teamSheet.get_as_df(
            include_tailing_empty=True, include_tailing_empty_rows=False
        )
        # display(Markdown("<h5 style='margin-left:4%;'>Select Project: </h5>"))
        projectSelected = widgets.Dropdown(
            options=self.teamSheetDataDF["PROJECT"].tolist(),
            value=None,
            description=" ",
            layout=self.get_centered_layout(),
        )
        projectSelected.observe(self.selectProject, names="value")
        with self.main_output:
            display(projectSelected)

    def selectProject(self, change):
        clear_output(wait=True)
        display(Markdown("<center><h5>Region: {}<h5></center> ".format(self.region)))
        display(
            Markdown("<center><h5>Team: {}<h5></center> ".format(self.teamSheet.title))
        )
        display(Markdown("<center><h5>Project: {}<h5></center> ".format(change.new)))
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

        # call levels function
        self.connectDCMS()
        self.setConstants()
        # if self.cloud.is_delivery_core_project:
        global all_levels_dict
        global ommiting_levels
        ommiting_levels = []
        all_levels_dict = self.gettingLevels_core()
        #display(all_levels_dict)
        # ssds
        self.get_ommiting_level_list(all_levels_dict)
        # else:
        #     all_levels = self.gettingLevels()
        # call levels function
        display(Markdown("<center><h5> Select levels to exclude: <h5></center> "))
        omit_levels = widgets.Dropdown(
            options=["Yes", "No"],
            description=" ",
            value=None,
            layout=self.get_centered_layout(),
        )
        omit_levels.observe(self.needdivision, names="value")
        with self.main_output:
            display(omit_levels)

    def needdivision(self, change):
        generate_button = widgets.Button(
            description="Generate",
            layout=self.get_centered_layout(display_width="20%"),
            button_style="info",
        )
        if change.new == "Yes":
            display(
                Markdown(
                    "<h5 style='margin-left:4%;'><center>Hold Ctrl/Cmd for multiple selection</center></h5>"
                )
            )
            self.omitted_levels = widgets.SelectMultiple(
                options=[""] + ommiting_levels,
                value=[""],
                description="Levels",
                disabled=False,
                layout=self.get_centered_layout(display_width="100%"),
            )
            # self.omitted_levels = widgets.Text(
            #     value="", description=" ", layout=self.get_centered_layout()
            # )
            # self.omitted_levels.on_submit(self.handle_submit)
            with self.main_output:
                display(self.omitted_levels)
                display(generate_button)
                # generate_button.on_click(self.handle_submit)
        else:
            with self.main_output:
                display(generate_button)
                # generate_button.on_click(self.handle_submit)
            # self.main_output.clear_output()
            # clear_output(wait=True)
            # self.lastRun()

        generate_button.on_click(self.handle_submit)

    def handle_submit(self, _):
        if self.omitted_levels is not None and len(self.omitted_levels.value) > 0:
            self.omitted_levels = self.omitted_levels.value
            # self.removing_ommited_levels()
        # self.main_output.clear_output()
        # clear_output(wait=True)
        # display("Janu!! I am handle")
        # self.lastRun()
        self.testMethods()

    def testMethods(self):
        self.setConstants()
        self.get_ommiting_level_list(all_levels_dict)
        self.removing_ommited_levels(all_levels_dict)

    def get_ommiting_level_list(self, all_levels_dict, parent_level=None):
        for level in all_levels_dict:
            if level != "~":
                ommiting_levels.append(
                    level if parent_level is None else parent_level + "<>" + level
                )
                if isinstance(all_levels_dict[level], list):
                    # ommiting_levels.append(level if parent_level is None else parent_level+"<>"+level)
                    for child in all_levels_dict[level]:
                        level_child = level + "<>" + child
                        if level_child not in ommiting_levels and child != "~":
                            ommiting_levels.append(level_child)
                else:
                    self.get_ommiting_level_list(all_levels_dict[level], level)

    def find_key(self, d):
        for key, value in d.items():
            yield [key, value]
            if isinstance(value, dict):
                yield from self.find_key(value)


    def get_parent_dictionary(self, d):
        for key, value in d.items():
            yield value
            if isinstance(value, dict):
                yield from self.get_parent_dictionary(value)

    def removing_ommited_levels(self, all_levels_dict):
        display(
            Markdown(
                "<center><h4>Removing ommitted levels from Dictionary</center></h4>"
            )
        )
        
        selected_omitted_levels = [x.split(",")[0] for x in self.omitted_levels]
        ommited_levels_keys = [x.split("<>") for x in selected_omitted_levels]
        #display('omitano', ommited_levels_keys)


        for x in ommited_levels_keys:
            for y in self.find_key(all_levels_dict):
                if y[0] == x[0]:
                    if isinstance(y[1], dict):
                        for z in self.get_parent_dictionary(all_levels_dict):
                            if isinstance(z, dict) and z.get(x[0]) is not None and z.get(x[0]).get(x[1]) is not None: 
                                z.get(x[0]).pop(x[1])
                                break
                            elif isinstance(z, dict) and z.get(x[0]) is None and z.get(x[1]) is not None:
                                z.pop(x[1])
                                break
                    elif isinstance(y[1], list):
                        y[1].remove(x[1])
 
 
        #uncomment the next line to show the results of the removing levels method
        display('IZBRISANO', all_levels_dict)

        # all_levels_dict = self.gettingLevels_core()
        # display(all_levels_dict)
        #final_level_Dataframe = all_levels_dict[
        #    ~(all_levels_dict.isin(map(lambda x: x.title().strip(), ll)).any(axis=1))
        #]
        # display(final_level_Dataframe)
        # asxx
        # newdftree = (
        #     newdftree.drop(["parent_0", "parent_1"], axis=1)
        #     .iloc[:, : self.num_level]
        #     .fillna(value="")
        # )

    # def level_setup(self):
    #     self.connectDCMS()
    #     self.setConstants()
    #     self.gettingLevels(time)

    def lastRun(self):
        self.setConstants()
        self.get_ommiting_level_list(all_levels_dict)
        self.removing_ommited_levels(all_levels_dict)
        self.get_all_sheet_references(self.auth)
        self.get_data_from_taskbreakdown()
        self.populatingLinkFigma()
        self.UpdatingSheet()

    def get_all_sheet_references(self, authorizing):
        clear_output(wait=True)
        display(
            Markdown("<center><h4>Getting information from the Sheets</center></h4>")
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

    def get_data_from_taskbreakdown(self):
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
        task_component_table = self.TBworkingSheetDF.loc[
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
        self.filteringData(task_component_table)
        # ! --------------------------------------------------------------------------------------- !

    def filteringData(self, data_to_filter):
        global datatoPaste
        display(data_to_filter)
        clear_output(wait=True)
        display(Markdown("<center><h4>Dropping Extra Information</center></h4>"))
        datatoPaste = data_to_filter.drop(
            [data_to_filter.columns.stop - 2, data_to_filter.columns.start + 2], axis=1
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

    # def transversingLevels(self, data, levelsData, parents):
    #     for levels in data:
    #         if levels["name"] != "Objects":
    #             parents.append(levels["name"] if "name" in levels else "#")
    #             if len(levels["children"]) > 0:
    #                 levelsData = self.transversingLevels(
    #                     levels["children"], levelsData, parents
    #                 )
    #             levelsData.append(
    #                 [
    #                     levels["name"] if "name" in levels else "",
    #                     levels["_id"],
    #                     levels["excluded"] if "excluded" in levels else "",
    #                     levels["template_type"] if "tempplate_type" in levels else "",
    #                     parents.copy(),
    #                 ]
    #             )
    #             parents.pop()
    #     return levelsData

    def converting_parents_to_list(self, parent_Data):
        for parent in parent_Data:
            if len(parent) > 0:
                self.dcmsius_stack_parent.append([x["name"] for x in parent])
            else:
                self.dcmsius_stack_parent.append(["#"])

    def gettingLevels_core(self):
        # allLevels = self.cloud.get_levels_tree()
        ALL_SPACE_IDS_QUERY = """
                query DCMSAllSpaces($customer: String!, $project: String!, $scope: Scope!){
                    spacesByFilter(tenant: { customer: $customer, project: $project, scope: $scope}, spaceFilters: { ancestorSpaceId: null }){
                        id
                        name
                        category
                        ancestors{
                            name
                        }
                    }
                }
                """
        allLevels = self.cloud.client._gql_client.query(ALL_SPACE_IDS_QUERY)
        all_level_dataframe = pd.DataFrame(allLevels["spacesByFilter"])
        all_level_dataframe["combined"] = all_level_dataframe["ancestors"]
        all_level_dataframe["ancestors"]
        for i, parents_list in all_level_dataframe.iterrows():
            parents_list["combined"].append({"name": parents_list["name"]})
            parents_list["combined"] = "<>".join(
                [x["name"] for x in parents_list["combined"]]
            )
        all_level_dataframe = all_level_dataframe[
            ["combined", "id", "name", "category", "ancestors"]
        ]
        self.converting_parents_to_list(all_level_dataframe["ancestors"])
        final_levels_df = pd.concat(
            [
                all_level_dataframe,
                pd.DataFrame(self.dcmsius_stack_parent).add_prefix("parent_"),
            ],
            axis=1,
        )
        final_levels_df.drop("ancestors", inplace=True, axis=1)

        # df_tree = pd.DataFrame(
        #     self.transversingLevels(allLevels["children"], [], ["#"]),
        #     columns=["name", "id", "excluded", "template_type", "ancestors"],
        # )
        # df_tree = pd.concat(
        #     [
        #         df_tree,
        #         pd.DataFrame(df_tree["ancestors"].values.tolist()).add_prefix(
        #             "parent_"
        #         ),
        #     ],
        #     axis=1,
        # )
        # newdftree = df_tree[[col for col in df_tree if col.startswith("parent")]]
        newdftree = self.filteringLevels(final_levels_df, self.num_level + 1)
        # if self.omitted_levels is not None:
        #     display(Markdown("<center><h4>Omitting levels</center></h4>"))
        #     newdftree = newdftree[
        #         ~(
        #             newdftree.isin(
        #                 map(lambda x: x.title().strip(), self.omitted_levels.split(","))
        #             ).any(axis=1)
        #         )
        #     ]
        # newdftree = (
        #     newdftree.drop(["parent_0", "parent_1"], axis=1)
        #     .iloc[:, : self.num_level]
        #     .fillna(value="")
        # )
        try:
            slicing_levels = newdftree.columns.get_loc(
                "parent_" + str(self.num_level + 1)
            )
        except KeyError:
            slicing_levels = len(newdftree.columns)
        newdftree = newdftree.iloc[:, :slicing_levels].fillna(value="~")
        # newdftree["combined_level"] = newdftree.agg(
        #     lambda row: "<>".join([x for x in row if x != ""]), axis=1
        # )
        newdftree = newdftree.drop_duplicates(subset=["combined"]).drop(
            ["combined", "category", "id", "name"], axis=1
        )
        # newdftree.to_csv("newdf.csv")
        for tab in self.upsTabNames:
            if (
                tab.lower() == "updatesheet"
                or tab.lower() == "update sheet"
                or tab is None
            ):
                level_final = newdftree.replace(to_replace="Objects", value="~")
            else:
                level_final = newdftree[newdftree.isin([tab]).any(axis=1)].replace(
                    to_replace="Objects", value="~"
                )
            LEVELSTOPASTE = {}
            for lst in level_final[level_final.columns].values:
                leaf = LEVELSTOPASTE
                for path in lst[:-2]:
                    if path != "~":
                        leaf = leaf.setdefault(path, {})
                if lst[-2] != "~":
                    if lst[-1] != "~":
                        leaf.setdefault(lst[-2], list()).append(lst[-1])
                    else:
                        leaf.setdefault(lst[-2], list()).append("~")
                else:
                    continue

            # for lst in level_final[level_final.columns].values:
            #     leaf = LEVELSTOPASTE
            #     for path in lst[:-2]:
            #         if path != "":
            #             leaf = leaf.setdefault(path, {})
            #     if lst[-2] != "" and lst[-1] != "":
            #         leaf.setdefault(lst[-2], list()).append(lst[-1])
        self.filtering_dictionary(LEVELSTOPASTE)
        return LEVELSTOPASTE
        # try:
        #     newSheet = self.updateSpreadSheet.add_worksheet(
        #         tab, rows=10000, cols=104
        #     )
        # except:
        #     self.updateSpreadSheet.del_worksheet(
        #         self.updateSpreadSheet.worksheet_by_title(tab)
        #     )
        #     newSheet = self.updateSpreadSheet.add_worksheet(
        #         tab, rows=10000, cols=104
        #     )
        # rowforlevels, columnforlevels = self.setConstants()
        # self.retransversing(
        #     LEVELSTOPASTE,
        #     rowforlevels,
        #     columnforlevels,
        #     newSheet.id,
        #     self.level_flag,
        #     {},
        #     {},
        #     {},
        #     {},
        #     self.level_count,
        #     0,
        # )

    def filtering_dictionary(self, levels_dict):
        for level_key in levels_dict:
            if not (bool(levels_dict[level_key])):
                levels_dict[level_key] = ["~"]
            elif (
                isinstance(levels_dict[level_key], list)
                and len(levels_dict[level_key]) > 1
            ):
                levels_dict[level_key].remove("~")
            elif isinstance(levels_dict[level_key], dict) and bool(
                levels_dict[level_key]
            ):
                self.filtering_dictionary(levels_dict[level_key])
            else:
                continue

    def filteringLevels(self, referenceLeveltreeData, totalParent):
        # if self.toplevelID[0] == "-" or self.toplevelID[0] == "":
        #     return referenceLeveltreeData[
        #         referenceLeveltreeData["parent_" + str(totalParent)] != ""
        #     ].reset_index(drop=True)
        # else:
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
                parentDict[Level] = [
                    rowforlevels,
                    columnforlevels,
                ]  # {"one Brooklane :["rownumebr", column],"floor 04":[r,c]"}
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
