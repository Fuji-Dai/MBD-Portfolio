classdef MTAS < matlab.apps.AppBase

    % Properties that correspond to app components
    properties (Access = public)
        MTASUIFigure                    matlab.ui.Figure
        ExecPanel                       matlab.ui.container.Panel
        tagetlistinfobookpathEditField  matlab.ui.control.EditField
        tagetlistinfobookpathEditFieldLabel  matlab.ui.control.Label
        tagetlistListBox                matlab.ui.control.ListBox
        tagetlistListBoxLabel           matlab.ui.control.Label
        targetmodelpathEditField        matlab.ui.control.EditField
        targetmodelpathEditFieldLabel   matlab.ui.control.Label
        GetInfobookButton               matlab.ui.control.Button
        AllexecButton                   matlab.ui.control.Button
        singleexecButton                matlab.ui.control.Button
        GettagetpathButton              matlab.ui.control.Button
        SettingPanel                    matlab.ui.container.Panel
        CreateFormat                    matlab.ui.control.CheckBox
        GetTestSpecBookButton           matlab.ui.control.Button
        genelatereportCheckBox          matlab.ui.control.CheckBox
        TestDataBeginCellTextArea       matlab.ui.control.TextArea
        TestDataBeginCellTextAreaLabel  matlab.ui.control.Label
        TimeDataBeginCellTextArea       matlab.ui.control.TextArea
        TimeDataBeginCellTextArea_2Label  matlab.ui.control.Label
        TestSpecIDBeginCellTextArea     matlab.ui.control.TextArea
        TestSpecIDBeginCellTextAreaLabel  matlab.ui.control.Label
        HeaderBeginCellTextArea         matlab.ui.control.TextArea
        HeaderBeginCellTextAreaLabel    matlab.ui.control.Label
        TestspecbookpathTextArea        matlab.ui.control.TextArea
        TestspecbookpathTextAreaLabel   matlab.ui.control.Label
        genelateresultCheckBox          matlab.ui.control.CheckBox
    end

    
    methods (Access = private)
        
    end
    

    % Callbacks that handle component events
    methods (Access = private)

        % Code that executes after component creation
        function startupFcn(app)
            % Getting the screen size
            screenSize = get(0, 'ScreenSize');   % [left bottom width height]
            % Get UIFigure Size
            figPos = app.MTASUIFigure.Position;      % [left bottom width height]
            % Calculation of the center position
            newLeft = (screenSize(3) - figPos(3)) / 2;
            newBottom = (screenSize(4) - figPos(4)) / 2;
            % Move the UIFigure to the center.
            app.MTASUIFigure.Position = [newLeft, newBottom, figPos(3), figPos(4)];

            fullpath = mfilename('fullpath');
            appFolder = fileparts(fullpath);
            if exist(fullfile(appFolder,'memoryData.mat'), 'file')
                memorydata = load(fullfile(appFolder,'memoryData.mat'));
                try
                    app.targetmodelpathEditField.Value = memorydata.targetfilepathEditField ;
                    app.tagetlistinfobookpathEditField.Value = memorydata.tagetlistinfobookpathEditField;
                    app.tagetlistListBox.Items = memorydata.tagetlistListBox;
                    app.TestspecbookpathTextArea.Value = memorydata.TestspecbookpathTextArea;
                catch

                end
            end
            app.genelatereportCheckBox.Enable = false;
        end

        % Button pushed function: GettagetpathButton
        function GettagetpathButtonPushed(app, event)
            try
                [file,location] = uigetfile('*.slx',...
                                   'Select model File(*.slx)', ...
                                   'MultiSelect', 'off');
                app.targetmodelpathEditField.Value = fullfile(location,file);
                figure(app.MTASUIFigure);
            catch
                % Do nothing
            end
        end

        % Button pushed function: GetInfobookButton
        function GetInfobookButtonPushed(app, event)
            try
                [file,location] = uigetfile('*.xlsx',...
                                   'Select excel file written target list(*.xlsx)', ...
                                   'MultiSelect', 'off');
                app.tagetlistinfobookpathEditField.Value = fullfile(location,file);
                tgtBookPath = app.tagetlistinfobookpathEditField.Value;
                tgtListInfo = readtable(tgtBookPath,...
                            'FileType','spreadsheet',...
                            'Sheet','tgtList',...
                            'ReadVariableNames',true);
                tgtList = {};
                for index = 1:numel(tgtListInfo.ScopeForTesting) 
                    if strcmp(char(tgtListInfo.ScopeForTesting(index)),'InScope')
                            tgtList = [tgtList ;fullfile(char(tgtListInfo.FoldaPath(index)),char(tgtListInfo.FileName(index)))];
                    end
                end
                app.tagetlistListBox.Items = tgtList;
                figure(app.MTASUIFigure);
            catch
                error("Portfolio version: execution disabled.");
            end
            figure(app.MTASUIFigure);
        end

        % Button pushed function: GetTestSpecBookButton
        function GetTestSpecBookButtonPushed(app, event)
            try
                [file,location] = uigetfile('*.xlsx',...
                                   'Select test spec book file(*.xlsx)', ...
                                   'MultiSelect', 'off');
                app.TestspecbookpathTextArea.Value = fullfile(location,file);
                
            catch
                
            end
            figure(app.MTASUIFigure);
        end

        % Button pushed function: singleexecButton
        function singleexecButtonPushed(app, event)
            try 
                curDirPath = pwd;
                tgtMdlPath = char(app.targetmodelpathEditField.Value);
                [tgtPath,tgtMdlName,~] = fileparts(tgtMdlPath);
                if ~isempty(tgtPath)
                    cd(tgtPath);
                end
                testSpecBookPath = char(app.TestspecbookpathTextArea.Value);
                listData = readtable(testSpecBookPath, 'FileType','spreadsheet','Sheet', 'List','ReadVariableNames',true);
                matchIdx = find(strcmp(listData.ModelName, tgtMdlName));
                testSheetName = char(listData.UnitID(matchIdx));
                genResultBook = true;
                genResultPDF = false;

                resultHeaderBgnCell = 'B1';
                resultDataBgnCell = 'B2';
                testSpecInHeaderBgnCell = char(app.HeaderBeginCellTextArea.Value);
                testSpecInDataBgnCell = char(app.TestDataBeginCellTextArea.Value);
                testSpecTimeDataBgnCell = char(app.TimeDataBeginCellTextArea.Value);
                testSpecIDBgnCell = char(app.TestSpecIDBeginCellTextArea.Value);
                msgObj = msgbox('proc processing...');

                if app.CreateFormat.Value
                    CreateTestSpecBook(testSheetName,tgtMdlPath,testSpecBookPath,testSpecInHeaderBgnCell,testSpecInDataBgnCell);
                elseif app.genelateresultCheckBox.Value
                    TestAutomation(tgtMdlPath,testSpecBookPath,testSheetName,genResultBook,genResultPDF,resultHeaderBgnCell,resultDataBgnCell,testSpecInHeaderBgnCell,testSpecInDataBgnCell,testSpecTimeDataBgnCell,testSpecIDBgnCell);
                end
                close(msgObj);
                cd(curDirPath);
                if app.CreateFormat.Value || app.genelateresultCheckBox.Value
                    msgbox('proc complete');
                end
            catch
                close(msgObj);
                errordlg("Portfolio version: execution disabled.");
            end
        end

        % Button pushed function: AllexecButton
        function AllexecButtonPushed(app, event)
            try
                curDirPath = pwd;
                testSpecBookPath = char(app.TestspecbookpathTextArea.Value);
                listData = readtable(testSpecBookPath, 'FileType','spreadsheet','Sheet', 'List','ReadVariableNames',true);

                genResultBook = true;
                genResultPDF = false;

                resultHeaderBgnCell = 'B1';
                resultDataBgnCell = 'B2';
                testSpecInHeaderBgnCell = char(app.HeaderBeginCellTextArea.Value);
                testSpecInDataBgnCell = char(app.TestDataBeginCellTextArea.Value);
                testSpecTimeDataBgnCell = char(app.TimeDataBeginCellTextArea.Value);
                testSpecIDBgnCell = char(app.TestSpecIDBeginCellTextArea.Value);

                tgtMdlList = app.tagetlistListBox.Items;
                progressObj = waitbar(0, 'Processing begin...');
                for i = 1:numel(tgtMdlList)
                    tgtMdlPath = char(tgtMdlList(i));
                    [tgtPath,tgtMdlName,~] = fileparts(tgtMdlPath);
                    if ~isempty(tgtPath)
                        cd(tgtPath);
                    end
                    matchIdx = find(strcmp(listData.ModelName, tgtMdlName));
                    testSheetName = char(listData.UnitID(matchIdx));
                    if app.CreateFormat.Value
                        CreateTestSpecBook(testSheetName,tgtMdlPath,testSpecBookPath,testSpecInHeaderBgnCell,testSpecInDataBgnCell);
                    elseif app.genelateresultCheckBox.Value
                        TestAutomation(tgtMdlPath,testSpecBookPath,testSheetName,genResultBook,genResultPDF,resultHeaderBgnCell,resultDataBgnCell,testSpecInHeaderBgnCell,testSpecInDataBgnCell,testSpecTimeDataBgnCell,testSpecIDBgnCell);
                    end
                    waitbar(i / numel(tgtMdlList), progressObj, sprintf('Processing... %d%%(%d/%d)', round(i / numel(tgtMdlList) * 100),i,numel(tgtMdlList)));
                end
                close(progressObj);
                cd(curDirPath);
                if app.CreateFormat.Value || app.genelateresultCheckBox.Value
                    msgbox('proc complete');
                end
            catch
                close(progressObj);
                errordlg("Portfolio version: execution disabled.");
            end
        end

        % Close request function: MTASUIFigure
        function MTASUIFigureCloseRequest(app, event)
            fullpath = mfilename('fullpath');
            appFolder = fileparts(fullpath);
            memorydata.targetfilepathEditField = app.targetmodelpathEditField.Value;
            memorydata.tagetlistinfobookpathEditField = app.tagetlistinfobookpathEditField.Value;
            memorydata.tagetlistListBox = app.tagetlistListBox.Items;
            memorydata.TestspecbookpathTextArea = app.TestspecbookpathTextArea.Value;
            save(fullfile(appFolder,'memoryData.mat'), '-struct', 'memorydata');
            delete(app)

        end

        % Value changed function: CreateFormat
        function CreateFormatValueChanged(app, event)
            if app.CreateFormat.Value
                app.genelateresultCheckBox.Value = false;
                app.genelateresultCheckBox.Enable = false;
            else
                app.genelateresultCheckBox.Enable = true;
            end
        end

        % Value changed function: genelateresultCheckBox
        function genelateresultCheckBoxValueChanged(app, event)
            if app.genelateresultCheckBox.Value
                app.CreateFormat.Value = false;
                app.CreateFormat.Enable = false;
                app.genelatereportCheckBox.Enable = true;
            else
                app.genelatereportCheckBox.Value = false;
                app.genelatereportCheckBox.Enable = false;
                app.CreateFormat.Enable = true;
            end            
        end
    end

    % Component initialization
    methods (Access = private)

        % Create UIFigure and components
        function createComponents(app)

            % Create MTASUIFigure and hide until all components are created
            app.MTASUIFigure = uifigure('Visible', 'off');
            app.MTASUIFigure.Position = [100 100 577 482];
            app.MTASUIFigure.Name = 'M-TAS';
            app.MTASUIFigure.CloseRequestFcn = createCallbackFcn(app, @MTASUIFigureCloseRequest, true);

            % Create SettingPanel
            app.SettingPanel = uipanel(app.MTASUIFigure);
            app.SettingPanel.Title = 'Setting';
            app.SettingPanel.Position = [5 309 569 166];

            % Create genelateresultCheckBox
            app.genelateresultCheckBox = uicheckbox(app.SettingPanel);
            app.genelateresultCheckBox.ValueChangedFcn = createCallbackFcn(app, @genelateresultCheckBoxValueChanged, true);
            app.genelateresultCheckBox.Text = 'genelate result';
            app.genelateresultCheckBox.Position = [430 42 101 22];

            % Create TestspecbookpathTextAreaLabel
            app.TestspecbookpathTextAreaLabel = uilabel(app.SettingPanel);
            app.TestspecbookpathTextAreaLabel.HorizontalAlignment = 'right';
            app.TestspecbookpathTextAreaLabel.Position = [13 108 112 22];
            app.TestspecbookpathTextAreaLabel.Text = 'Test spec book path';

            % Create TestspecbookpathTextArea
            app.TestspecbookpathTextArea = uitextarea(app.SettingPanel);
            app.TestspecbookpathTextArea.Position = [140 108 255 24];

            % Create HeaderBeginCellTextAreaLabel
            app.HeaderBeginCellTextAreaLabel = uilabel(app.SettingPanel);
            app.HeaderBeginCellTextAreaLabel.HorizontalAlignment = 'right';
            app.HeaderBeginCellTextAreaLabel.Position = [15 65 102 22];
            app.HeaderBeginCellTextAreaLabel.Text = 'Header Begin Cell';

            % Create HeaderBeginCellTextArea
            app.HeaderBeginCellTextArea = uitextarea(app.SettingPanel);
            app.HeaderBeginCellTextArea.Position = [140 63 54 27];
            app.HeaderBeginCellTextArea.Value = {'G8'};

            % Create TestSpecIDBeginCellTextAreaLabel
            app.TestSpecIDBeginCellTextAreaLabel = uilabel(app.SettingPanel);
            app.TestSpecIDBeginCellTextAreaLabel.HorizontalAlignment = 'right';
            app.TestSpecIDBeginCellTextAreaLabel.Position = [203 16 131 22];
            app.TestSpecIDBeginCellTextAreaLabel.Text = 'Test Spec ID Begin Cell';

            % Create TestSpecIDBeginCellTextArea
            app.TestSpecIDBeginCellTextArea = uitextarea(app.SettingPanel);
            app.TestSpecIDBeginCellTextArea.Position = [341 14 54 27];
            app.TestSpecIDBeginCellTextArea.Value = {'A11'};

            % Create TimeDataBeginCellTextArea_2Label
            app.TimeDataBeginCellTextArea_2Label = uilabel(app.SettingPanel);
            app.TimeDataBeginCellTextArea_2Label.HorizontalAlignment = 'right';
            app.TimeDataBeginCellTextArea_2Label.Position = [13 16 121 22];
            app.TimeDataBeginCellTextArea_2Label.Text = 'Time Data  Begin Cell';

            % Create TimeDataBeginCellTextArea
            app.TimeDataBeginCellTextArea = uitextarea(app.SettingPanel);
            app.TimeDataBeginCellTextArea.Position = [140 14 54 27];
            app.TimeDataBeginCellTextArea.Value = {'E11'};

            % Create TestDataBeginCellTextAreaLabel
            app.TestDataBeginCellTextAreaLabel = uilabel(app.SettingPanel);
            app.TestDataBeginCellTextAreaLabel.HorizontalAlignment = 'right';
            app.TestDataBeginCellTextAreaLabel.Position = [203 65 114 22];
            app.TestDataBeginCellTextAreaLabel.Text = 'Test Data Begin Cell';

            % Create TestDataBeginCellTextArea
            app.TestDataBeginCellTextArea = uitextarea(app.SettingPanel);
            app.TestDataBeginCellTextArea.Position = [341 63 54 27];
            app.TestDataBeginCellTextArea.Value = {'G11'};

            % Create genelatereportCheckBox
            app.genelatereportCheckBox = uicheckbox(app.SettingPanel);
            app.genelatereportCheckBox.Text = 'genelate report';
            app.genelatereportCheckBox.Position = [430 14 103 22];

            % Create GetTestSpecBookButton
            app.GetTestSpecBookButton = uibutton(app.SettingPanel, 'push');
            app.GetTestSpecBookButton.ButtonPushedFcn = createCallbackFcn(app, @GetTestSpecBookButtonPushed, true);
            app.GetTestSpecBookButton.Position = [422 109 121 23];
            app.GetTestSpecBookButton.Text = 'Get Test Spec Book';

            % Create CreateFormat
            app.CreateFormat = uicheckbox(app.SettingPanel);
            app.CreateFormat.ValueChangedFcn = createCallbackFcn(app, @CreateFormatValueChanged, true);
            app.CreateFormat.Text = 'create format';
            app.CreateFormat.Position = [430 70 93 22];

            % Create ExecPanel
            app.ExecPanel = uipanel(app.MTASUIFigure);
            app.ExecPanel.Title = 'Exec';
            app.ExecPanel.Position = [6 9 569 288];

            % Create GettagetpathButton
            app.GettagetpathButton = uibutton(app.ExecPanel, 'push');
            app.GettagetpathButton.ButtonPushedFcn = createCallbackFcn(app, @GettagetpathButtonPushed, true);
            app.GettagetpathButton.Position = [385 234 86 23];
            app.GettagetpathButton.Text = 'Get taget path';

            % Create singleexecButton
            app.singleexecButton = uibutton(app.ExecPanel, 'push');
            app.singleexecButton.ButtonPushedFcn = createCallbackFcn(app, @singleexecButtonPushed, true);
            app.singleexecButton.Position = [478 234 77 23];
            app.singleexecButton.Text = 'single exec';

            % Create AllexecButton
            app.AllexecButton = uibutton(app.ExecPanel, 'push');
            app.AllexecButton.ButtonPushedFcn = createCallbackFcn(app, @AllexecButtonPushed, true);
            app.AllexecButton.Position = [479 191 77 23];
            app.AllexecButton.Text = 'All exec';

            % Create GetInfobookButton
            app.GetInfobookButton = uibutton(app.ExecPanel, 'push');
            app.GetInfobookButton.ButtonPushedFcn = createCallbackFcn(app, @GetInfobookButtonPushed, true);
            app.GetInfobookButton.Position = [385 191 86 23];
            app.GetInfobookButton.Text = 'Get Info book';

            % Create targetmodelpathEditFieldLabel
            app.targetmodelpathEditFieldLabel = uilabel(app.ExecPanel);
            app.targetmodelpathEditFieldLabel.HorizontalAlignment = 'right';
            app.targetmodelpathEditFieldLabel.Position = [36 234 98 22];
            app.targetmodelpathEditFieldLabel.Text = 'target model path';

            % Create targetmodelpathEditField
            app.targetmodelpathEditField = uieditfield(app.ExecPanel, 'text');
            app.targetmodelpathEditField.Position = [149 234 222 22];

            % Create tagetlistListBoxLabel
            app.tagetlistListBoxLabel = uilabel(app.ExecPanel);
            app.tagetlistListBoxLabel.HorizontalAlignment = 'right';
            app.tagetlistListBoxLabel.Position = [86 141 50 22];
            app.tagetlistListBoxLabel.Text = 'taget list';

            % Create tagetlistListBox
            app.tagetlistListBox = uilistbox(app.ExecPanel);
            app.tagetlistListBox.Items = {'C:\'};
            app.tagetlistListBox.Position = [149 10 406 153];
            app.tagetlistListBox.Value = 'C:\';

            % Create tagetlistinfobookpathEditFieldLabel
            app.tagetlistinfobookpathEditFieldLabel = uilabel(app.ExecPanel);
            app.tagetlistinfobookpathEditFieldLabel.HorizontalAlignment = 'right';
            app.tagetlistinfobookpathEditFieldLabel.Position = [7 191 128 22];
            app.tagetlistinfobookpathEditFieldLabel.Text = 'taget list info book path';

            % Create tagetlistinfobookpathEditField
            app.tagetlistinfobookpathEditField = uieditfield(app.ExecPanel, 'text');
            app.tagetlistinfobookpathEditField.HorizontalAlignment = 'right';
            app.tagetlistinfobookpathEditField.Position = [149 191 222 22];

            % Show the figure after all components are created
            app.MTASUIFigure.Visible = 'on';
        end
    end

    % App creation and deletion
    methods (Access = public)

        % Construct app
        function app = MTAS

            % Create UIFigure and components
            createComponents(app)

            % Register the app with App Designer
            registerApp(app, app.MTASUIFigure)

            % Execute the startup function
            runStartupFcn(app, @startupFcn)

            if nargout == 0
                clear app
            end
        end

        % Code that executes before app deletion
        function delete(app)

            % Delete UIFigure when app is deleted
            delete(app.MTASUIFigure)
        end
    end
end