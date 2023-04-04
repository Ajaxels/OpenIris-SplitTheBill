classdef Controller < handle
    % @type Controller class is a template class for using with
    % GUI developed using appdesigner of Matlab
    
	% Copyright (C) 2019-2020 Ilya Belevich, University of Helsinki (ilya.belevich @ helsinki.fi)
    % The MIT License (https://opensource.org/licenses/MIT)
    
    
    properties
        Model
        % handles to the model
        View
        % handle to the view
        listener
        % a cell array with handles to listeners
        childControllers
        % list of opened subcontrollers
        childControllersIds
        % a cell array with names of initialized child controllers
    end
    
    events
        %> Description of events
        closeEvent
        % event firing when window is closed
    end
    
    methods (Static)
        function purgeControllers(obj, src, evnt)
            % find index of the child controller
            id = obj.findChildId(class(src));
            
            % delete the child controller
            delete(obj.childControllers{id});
            
            % clear the handle
            obj.childControllers(id) = [];
            obj.childControllersIds(id) = [];
        end
        
        
        function ViewListner_Callback(obj, src, evnt)
            switch evnt.EventName
                case {'updateGuiWidgets'}
                    obj.updateWidgets();
            end
        end
    end
    
    methods
        % declaration of functions in the external files, keep empty line in between for the doc generator
        id = findChildId(obj, childName)        % find index of a child controller  
        
        startController(obj, controllerName, varargin)        % start a child controller
        
        function obj = Controller(Model, parameter)
            obj.Model = Model;    % assign model
            obj.View = View(obj);
            
            obj.View.gui.Name = [obj.View.gui.Name, '  ' parameter];
            
            % obtain settings from a file
            % saving settings
            temp = tempdir;
            if exist(fullfile(temp, 'split_bills_settings.mat'), 'file') == 2
                load(fullfile(temp, 'split_bills_settings.mat'));
                obj.Model.Settings = mibConcatenateStructures(obj.Model.Settings, Settings);    % concatenate Settings structure
                fprintf('Loading settings from %s\n', fullfile(temp, 'split_bills_settings.mat'));
            end
            
            obj.updateWidgets();
			
			% add listner to obj.mibModel and call controller function as a callback
            % option 1: recommended, detects event triggered by mibController.updateGuiWidgets
            obj.listener{1} = addlistener(obj.Model, 'updateGuiWidgets', @(src,evnt) obj.ViewListner_Callback(obj, src, evnt));    % listen changes in number of ROIs
        end
        
        function closeWindow(obj)
            % closing Controller window
            if isvalid(obj.View.gui)
                delete(obj.View.gui);   % delete childController window
            end
            
            % saving settings
            temp = tempdir;
            Settings = obj.Model.Settings;
            save(fullfile(temp, 'split_bills_settings.mat'), 'Settings');
            
            % delete listeners, otherwise they stay after deleting of the
            % controller
            for i=1:numel(obj.listener)
                delete(obj.listener{i});
            end
            
            notify(obj, 'closeEvent');      % notify mibController that this child window is closed
        end
        
        function updateFilename(obj, value)
            % update input filename
            if exist(value, 'file') == 2
                obj.Model.Settings.gui.InputFilename = value;
                obj.Model.getColumnNames();
                if ismember(obj.View.TableIndexField.Value, obj.Model.VariableNames)
                    obj.View.TableIndexField.Items = sort(obj.Model.VariableNames);
                else
                    obj.View.TableIndexField.Items = sort(obj.Model.VariableNames);
                    if ismember('ID', obj.Model.VariableNames)
                        obj.View.TableIndexField.Value = 'ID';
                    end
                end

                if ismember(obj.View.SplitBillsField.Value, obj.Model.VariableNames)
                    obj.View.SplitBillsField.Items = sort(obj.Model.VariableNames);
                else
                    obj.View.SplitBillsField.Items = sort(obj.Model.VariableNames);
                    if ismember('RequestID', obj.Model.VariableNames)
                        obj.View.SplitBillsField.Value = 'RequestID';
                    end
                end

                if ismember(obj.View.SortBillsField.Value, obj.Model.VariableNames)
                    obj.View.SortBillsField.Items = sort(obj.Model.VariableNames);
                else
                    obj.View.SortBillsField.Items = sort(obj.Model.VariableNames);
                    if ismember('BookingStart', obj.Model.VariableNames)
                        obj.View.SortBillsField.Value = 'BookingStart';
                    end
                end
                
                obj.View.UITable.ColumnName = obj.Model.VariableNames;
                obj.View.UITable.Data = obj.Model.T(1,:);
            else
                obj.View.InputFilename.Value = obj.Model.Settings.gui.InputFilename;
            end
        end
        
        function updateDirectory(obj, value)
            % update output directory name
            if exist(value, 'dir') ~= 7     
                errordlg(sprintf('!!! Error !!!\n\nDirectory with name:\n%s\ndoes not exist!', value), 'Wrong output directory');
                if exist(obj.Model.Settings.gui.OutputDirectory, 'dir') == 7
                    obj.View.OutputDirectory.Value = obj.Model.Settings.gui.OutputDirectory;
                else
                    obj.View.OutputDirectory.Value = pwd;
                end
            else
                obj.Model.Settings.gui.OutputDirectory =  value;
            end
        end
        
        function updateWidgets(obj)
            % function updateWidgets(obj)
            % update widgets of this window
            
            fieldNames = fieldnames(obj.Model.Settings.gui);
            for fieldId = 1:numel(fieldNames)
                if ~isprop(obj.View, fieldNames{fieldId}); continue; end
                switch obj.View.(fieldNames{fieldId}).Type
                    case 'uidropdown'
                        obj.View.(fieldNames{fieldId}).Items = sort({obj.Model.Settings.gui.(fieldNames{fieldId})});
                        obj.View.(fieldNames{fieldId}).Value = obj.Model.Settings.gui.(fieldNames{fieldId});
                    otherwise
                        obj.View.(fieldNames{fieldId}).Value = obj.Model.Settings.gui.(fieldNames{fieldId});
                end
            end
        end
        
        function updateSettings(obj, event)
%             switch event.Source.Type    % event.Source == obj.View.SplitBillsField
%                 case 'uidropdown'
%                     if ~isempty(event.Source.Value)
%                     
%                     end
%                 otherwise
%                     if ~isempty(event.Source.Value)
%                         obj.Model.Settings.gui.InputFilename
%                     end
%             end
                
            % update program settings from the widgets
            if ~isempty(obj.View.InputFilename.Value)
                obj.Model.Settings.gui.InputFilename = obj.View.InputFilename.Value; 
            end
            if ~isempty(obj.View.SplitBillsField.Items)
                obj.Model.Settings.gui.SplitBillsField = obj.View.SplitBillsField.Value;
                if ~isempty(obj.Model.T)
                    if isnumeric(obj.Model.T.(obj.View.SplitBillsField.Value))
                        obj.View.FieldExampleText.Text = num2str(obj.Model.T.(obj.View.SplitBillsField.Value)(1));
                        obj.View.FieldExampleText.Tooltip = num2str(obj.Model.T.(obj.View.SplitBillsField.Value)(1));
                    else
                        obj.View.FieldExampleText.Text = obj.Model.T.(obj.View.SplitBillsField.Value)(1);
                        obj.View.FieldExampleText.Tooltip = obj.Model.T.(obj.View.SplitBillsField.Value)(1);
                    end
                end
            end
            if ~isempty(obj.View.TableIndexField.Items)
                obj.Model.Settings.gui.TableIndexField = obj.View.TableIndexField.Value;
            end
            if ~isempty(obj.View.SplitBillsField.Items)
                obj.Model.Settings.gui.SortBillsField = obj.View.SortBillsField.Value;
            end
            if ~isempty(obj.View.ResponsiblePerson.Value)
                obj.Model.Settings.gui.ResponsiblePerson = obj.View.ResponsiblePerson.Value;
            end
            if ~isempty(obj.View.ProviderName.Value)
                obj.Model.Settings.gui.ProviderName = obj.View.ProviderName.Value;
            end
            if ~isempty(obj.View.OutputDirectory.Value)
                obj.Model.Settings.gui.OutputDirectory = obj.View.OutputDirectory.Value; 
            end
            obj.Model.Settings.gui.GenerateSummaryFile = obj.View.GenerateSummaryFile.Value; 
            if ~isempty(obj.View.HeaderStartingCell.Value)
                obj.Model.Settings.gui.HeaderStartingCell = obj.View.HeaderStartingCell.Value; 
            end
            if ~isempty(obj.View.DataStartingCell.Value)
                obj.Model.Settings.gui.DataStartingCell = obj.View.DataStartingCell.Value; 
            end
            obj.Model.Settings.gui.DetectDuplicates = obj.View.DetectDuplicates.Value;
            obj.Model.Settings.gui.DetectCollaborations = obj.View.DetectCollaborations.Value;
            obj.Model.Settings.gui.CollaborationMarker = obj.View.CollaborationMarker.Value;
            obj.Model.Settings.gui.DetectExternalProjects = obj.View.DetectExternalProjects.Value;
            obj.Model.Settings.gui.ExternalProjectMarker = obj.View.ExternalProjectMarker.Value;
        end
        
        function StartProcessing(obj)
            % start processing
            obj.View.StartProcessingButton.BackgroundColor = 'r';
            obj.Model.start();
            obj.View.StartProcessingButton.BackgroundColor = 'g';
        end
    end
end