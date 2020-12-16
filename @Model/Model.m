classdef Model < handle
    % classdef Model < handle
    % the main model class of SplitTheBill
    
    % Copyright (C) 2019-2020 Ilya Belevich, University of Helsinki (ilya.belevich @ helsinki.fi)
    % The MIT License (https://opensource.org/licenses/MIT)
    
    properties 
        T
        % table with imported excel sheet
        Settings
        % a structure with settings
        % .InputFilename - name of the input file
        % .OutputDirectory - output directory
        % .SplitBillsField - field used to split the bills
        VariableNames
        % list of column names from excel sheet
    end
    
    properties (SetObservable)

    end
        
    events
        updateGuiWidgets
        % event after update of GuiWidgets of Controller
    end
    
    methods
        % declaration of functions in the external files, keep empty line in between for the doc generator
        
        BatchOptOut = selectFile(obj, BatchOpt)   % choose a file
        
        function obj = Model()
            obj.reset();
        end
        
        function reset(obj)
            obj.T = [];  
            obj.Settings = struct();     % current session settings
            obj.Settings.gui = struct();     % current settings for gui widgets
            obj.Settings.gui.InputFilename = pwd;
            obj.Settings.gui.OutputDirectory = pwd;
            obj.Settings.gui.TableIndexField = 'ID';   % name of the field with index, for detection of duplicates
            obj.Settings.gui.SplitBillsField = '';
            obj.Settings.gui.SortBillsField = '';
            obj.Settings.gui.DetectDuplicates = true;
            obj.Settings.gui.ResponsiblePerson = '';
            obj.Settings.gui.ProviderName = '';
            obj.Settings.gui.GenerateSummaryFile = true;
            obj.Settings.gui.HeaderStartingCell = 'A3';
            obj.Settings.gui.DataStartingCell = 'A4';
            obj.Settings.gui.DetectCollaborations = true;
            obj.Settings.gui.CollaborationMarker = '[C]';
        end
        
        function getColumnNames(obj)
            % get column names from the excel file
            wb = waitbar(0, sprintf('Obtaining names for the column\nPlease wait...'));
            warning('off', 'MATLAB:table:ModifiedAndSavedVarnames');
            waitbar(0.05, wb);
            rangeText = sprintf('%s:%s', obj.Settings.gui.HeaderStartingCell(2:end), obj.Settings.gui.DataStartingCell(2:end));
            obj.T = readtable(obj.Settings.gui.InputFilename, 'Range', rangeText);   % read excel file
            waitbar(0.9, wb);
            obj.VariableNames = obj.T.Properties.VariableNames;
            waitbar(1, wb);
            delete(wb);
        end
        
        function start(obj)
            tic
            warning('off', 'MATLAB:xlswrite:AddSheet');
            wb = waitbar(0, sprintf('Splitting the bills\nPlease wait...'));
            
            opts = detectImportOptions(obj.Settings.gui.InputFilename, 'NumHeaderLines', 0);
            opts.VariableNamesRange = obj.Settings.gui.HeaderStartingCell;
            opts.DataRange = obj.Settings.gui.DataStartingCell;
            obj.T = readtable(obj.Settings.gui.InputFilename, opts, 'ReadVariableNames', true);   % read excel file
            obj.VariableNames = obj.T.Properties.VariableNames;
            
            tableIndex = find(ismember(obj.VariableNames, obj.Settings.gui.TableIndexField));
            splitIndex = find(ismember(obj.VariableNames, obj.Settings.gui.SplitBillsField));
            sortingIndex = find(ismember(obj.VariableNames, obj.Settings.gui.SortBillsField));
            
            CreationDateIndex = find(ismember(obj.VariableNames, 'CreationDate'));
            GroupPIIndex = find(ismember(obj.VariableNames, 'Group'));
            ProjectNameIndex = find(ismember(obj.VariableNames, 'RequestTitle'));
            RequestIDIndex = find(ismember(obj.VariableNames, 'RequestID'));
            AffiliatedIndex = find(ismember(obj.VariableNames, 'Affiliated department'));   % missing in the excel sheet
            OrganizationIndex = find(ismember(obj.VariableNames, 'Organization'));  
            CostCenterNameIndex = find(ismember(obj.VariableNames, 'CostCenterName')); 
            CostCenterCodeIndex = find(ismember(obj.VariableNames, 'CostCenterCode')); 
            RemitCodeIndex = find(ismember(obj.VariableNames, 'RemitCode')); 
            PriceTypeIndex = find(ismember(obj.VariableNames, 'PriceType')); 
            ChargeIndex = find(ismember(obj.VariableNames, 'Charge'));
            ResourceIndex = find(ismember(obj.VariableNames, 'Resource'));
            ChargeTypeIndex = find(ismember(obj.VariableNames, 'ChargeType'));
            UserNameIndex = find(ismember(obj.VariableNames, 'UserName'));
            DescriptionIndex = find(ismember(obj.VariableNames, 'Description'));
            BillingAddressIndex = find(ismember(obj.VariableNames, 'BillingAddress'));
            
            QuantityIndex = find(ismember(obj.VariableNames, 'Quantity'));
            ProductIndex = find(ismember(obj.VariableNames, 'PriceList_Product'));
            splitEntries = table2cell(unique(obj.T(:, splitIndex)));
            dateString = datestr(date, 'yymmdd');
            
            % detect if there are some fields that do not have any values
            % in the splitting field
            missingSplitFieldValues = find(cellfun(@isempty, splitEntries)); %#ok<EFIND>
            if ~isempty(missingSplitFieldValues)
                warndlg(sprintf('!!! Warning !!!\n\nPlease check the "%s" field in the original Excel file!\nIt looks that some of those fields are empty.\nTo proceed further all fields used for splitting should contain some data!', obj.VariableNames{splitIndex}), ...
                    'Missing info');
                delete(wb);
                return;
            end
            
            if obj.Settings.gui.GenerateSummaryFile
                Summary = cell([numel(splitEntries)+1, 6]);     % allocate space
                Summary{1,1} = 'Group name'; 
                %Summary{1,2} = 'Affiliated department'; 
                Summary{1,2} = 'Organization'; 
                Summary{1,3} = 'Remit code'; 
                Summary{1,4} = 'Cost center code'; 
                Summary{1,5} = 'Total charge'; 
                Summary{1,6} = 'Price type'; 
                Summary{1,7} = 'Request title'; 
                SummaryCounter = 2;     % row index in the summary file, 1st one reserved for the titles
            end
            
            if obj.Settings.gui.DetectCollaborations
                collaborationsDir = fullfile(obj.Settings.gui.OutputDirectory, 'Collaborations');
                if ~isfolder(collaborationsDir); mkdir(collaborationsDir); end
            end
            
            % look for duplicate indices
            if obj.Settings.gui.DetectDuplicates
                T2 = sortrows(obj.T(:, tableIndex));
                [~, indices] = unique(T2);
                duplicateIndices = setdiff(1:numel(T2), indices);
                if ~isempty(duplicateIndices)
                    duplicateValues = T2.(obj.VariableNames{tableIndex})(duplicateIndices);
                    delete(wb);
                    errordlg(sprintf('!!! Error !!!\n\nThe duplicate indices were found in the table, remove the duplicates in OpenIris and regenerate the invoice file!\nList of duplicate indices:\n%s', strjoin(duplicateValues, ', ')), ...
                        'Duplicate indices!');
                    return;
                end
            end
            
            for splitId=1:numel(splitEntries)
                indices = ismember(table2cell(obj.T(:, splitIndex)), splitEntries{splitId});
                T2 = obj.T(indices, :);
                T2.Properties.VariableNames = obj.VariableNames;
                T2 = sortrows(T2, sortingIndex);
                fnTemplate = sprintf('%s_%s_%s.xlsx', cell2mat(table2array(T2(1, GroupPIIndex))), cell2mat(table2array(T2(1, splitIndex))), dateString); %#ok<FNDSB>
                
                fn = fullfile(obj.Settings.gui.OutputDirectory, fnTemplate); % default output filename
                if obj.Settings.gui.DetectCollaborations    % detect collaboration project to put them into a separate folder
                    if ~isempty(strfind(T2.RequestTitle{1}, obj.Settings.gui.CollaborationMarker))  % find the text marker
                        fn = fullfile(collaborationsDir, fnTemplate); % update output filename for collaboration projects
                    end
                end
                if exist(fn, 'file') == 2; delete(fn); end
                
                % do calculations for summary
                % calculate starting and ending dates
                startingDate = min(datetime(table2cell(T2(:, CreationDateIndex)), 'InputFormat','yy-MM-dd HH:mm'));
                startingDate = datestr(startingDate, 'dd.mm.yyyy');
                endingDate = max(datetime(table2cell(T2(:, CreationDateIndex)), 'InputFormat','yy-MM-dd HH:mm'));
                endingDate = datestr(endingDate, 'dd.mm.yyyy');
                [~, invoiceName] = fileparts(fnTemplate);
                
                % get indices of the products
                productIndices = ismember(table2array(T2(:, ChargeTypeIndex)), 'Product (request)');
                % generate table with products
                productsTable = T2(productIndices, :);
                reservationsTable = T2(~productIndices, :);
                
                %T2(productIndices,:) = [];
                resourceList = table2array(unique(reservationsTable(:, ResourceIndex)));
                
                clear s;
                % generate the invoice
                s{1,3} = 'Instrument and product invoice';
                s{3,1} = 'Time period of invoice:'; s{3,3} = [startingDate ' - ' endingDate];
                s{4,1} = 'Invoice name:'; s{4,3} = invoiceName;
                s{5,1} = 'Responsible person:'; s{5,3} = obj.Settings.gui.ResponsiblePerson;
                s{6,1} = 'Provider:'; s{6,3} = obj.Settings.gui.ProviderName;
                
                s{3,4} = 'Billing address:';
                s{3,5} = cell2mat(table2cell(T2(1, BillingAddressIndex)));
                
                s{8,1} = 'Project title:'; s{8,3} = [cell2mat(table2cell(T2(1, RequestIDIndex))) ', ' cell2mat(table2cell(T2(1, ProjectNameIndex)))]; %#ok<*FNDSB>
                s{10,1} = 'Group leader:'; s{10,3} = cell2mat(table2cell(T2(1, GroupPIIndex))); 
                %s{10,1} = 'Affiliated department:'; s{10,3} = ''; 
                s{11,1} = 'Organization:'; s{11,3} = cell2mat(table2cell(T2(1, OrganizationIndex)));
                
                s{13,3} = 'Cost center name'; s{13,4} = 'Remit code'; s{13,5} = 'Cost center code'; s{13,6} = 'Price type'; 
                s{14,1} = 'Cost center:'; 
                s{14,3} = cell2mat(table2cell(T2(1, CostCenterNameIndex)));
                s{14,4} = cell2mat(table2cell(T2(1, RemitCodeIndex)));
                s{14,5} = cell2mat(table2cell(T2(1, CostCenterCodeIndex)));
                if ~isempty(reservationsTable)  % fetch price type from the reservations list
                    PriceTypeVar = cell2mat(table2cell(reservationsTable(1, PriceTypeIndex)));
                else    % fetch price type from the products, but do not take Default if any other type is present
                    PriceTypeList = table2cell(productsTable(:, PriceTypeIndex));
                    PriceTypeListIndex = find(~ismember(PriceTypeList, 'Default') == 1, 1);
                    if ~isempty(PriceTypeListIndex)
                        PriceTypeVar = PriceTypeList{PriceTypeListIndex};
                    else
                        PriceTypeVar = cell2mat(table2cell(productsTable(1, PriceTypeIndex)));
                    end
                end
                s{14,6} = PriceTypeVar;
                
                shiftY = 17;
                if sum(productIndices) > 0
                    % products exist
                    lineVec = zeros([numel(resourceList)+2, 1]);  % where to draw a bottom border
                    productsSwitch = 1;
                else
                    % no products
                    lineVec = zeros([numel(resourceList)+1, 1]);  % where to draw a bottom border
                    productsSwitch = 0;
                end
                
                for resId = 1:numel(resourceList)
                    s{shiftY, 1} = resourceList{resId};
                    s{shiftY, 3} = 'Researcher name';
                    s{shiftY, 4} = 'Reserved hours (h.m.s)';
                    s{shiftY, 5} = 'Price, euros';
                    s{shiftY, 6} = 'Price type';
                    lineVec(resId) = shiftY;
                    
                    % get reservation list for the resource
                    indices = ismember(table2cell(reservationsTable(:, ResourceIndex)), resourceList{resId});
                    T3 = reservationsTable(indices, :);
                    userList = table2array(unique(T3(:, UserNameIndex)));
                    for userId = 1:numel(userList)
                        indices2 = ismember(table2cell(T3(:, UserNameIndex)), userList{userId});
                        Tuser = T3(indices2, :);
                        descVec = table2cell(Tuser(:, DescriptionIndex));
                        clip1 = arrayfun(@(x) strfind(x, ' to'), descVec);
                        clip2 = arrayfun(@(x) strfind(x, ', resource'), descVec);
                        resTime = 0;
                        for i=1:size(descVec, 1)
                            startTime = descVec{i}(1:clip1{i}-1);
                            endTime = descVec{i}(clip1{i}+4:clip2{i}-1);
                            startTime = datetime(startTime, 'InputFormat','yyyy-MM-dd HH:mm');
                            endTime = datetime(endTime, 'InputFormat','yyyy-MM-dd HH:mm');
                            resTime = resTime + endTime - startTime;
                        end
                        shiftY = shiftY + 1;
                        s{shiftY,3} = cell2mat(table2cell(Tuser(1, UserNameIndex)));
                        s{shiftY,4} = sprintf('%s', resTime);
                        Tuser2 = Tuser(:, ChargeIndex);    Tuser2.Charge = str2double(Tuser2.Charge);
                        s{shiftY,5} = sprintf('%0.2f', sum(Tuser2.Charge));
                        s{shiftY,6} = cell2mat(table2cell(Tuser(1, PriceTypeIndex)));
                    end
                    shiftY = shiftY + 2;
                end
                
                % add products
                if productsSwitch
                    if isempty(resId);  resId = 0; end  % when no reservations
                    resId = resId + 1;
                    s{shiftY, 1} = 'Products';
                    s{shiftY, 3} = 'Description';
                    s{shiftY, 4} = 'Quantity';
                    s{shiftY, 5} = 'Price, euros';
                    s{shiftY, 6} = 'Price type';
                    lineVec(resId) = shiftY;    % index of the line where to add underline
                    
                    % get reservation list for the resource
                    productList = table2array(unique(productsTable(:, ProductIndex)));
                    for productId = 1:numel(productList)
                        indices2 = ismember(table2cell(productsTable(:, ProductIndex)), productList{productId});
                        CurrProduct = productsTable(indices2, :);
                        
                        shiftY = shiftY + 1;
                        s{shiftY,3} = cell2mat(table2cell(CurrProduct(1, ProductIndex)));
                        quantityVec = CurrProduct(:, QuantityIndex);
                        s{shiftY,4} = sum(str2double(quantityVec.Quantity));
                        
                        chargeVec = CurrProduct(:, ChargeIndex);    chargeVec.Charge = str2double(chargeVec.Charge);
                        s{shiftY,5} = sprintf('%0.2f', sum(chargeVec.Charge));
                        s{shiftY,6} = cell2mat(table2cell(CurrProduct(1, PriceTypeIndex)));
                    end
                    shiftY = shiftY + 2;
                end
                
                s{shiftY,1} = 'Summary:'; s{shiftY,3} = 'Group name'; 
                %s{shiftY,4} = 'Affiliated department'; 
                s{shiftY,4} = 'Remit code'; s{shiftY,5} = 'Cost center code'; s{shiftY,6} = 'Total charge'; %s{shiftY,8} = 'Price type'; 
                resId = resId + 1;
                lineVec(resId) = shiftY;    % index of the line where to add underline
                shiftY = shiftY + 1;
                s{shiftY,3} = cell2mat(table2cell(T2(1, GroupPIIndex))); 
                %s{shiftY,4} = ''; 
                s{shiftY,4} = cell2mat(table2cell(T2(1, RemitCodeIndex))); 
                s{shiftY,5} = cell2mat(table2cell(T2(1, CostCenterCodeIndex)));
                T3 = T2(:, ChargeIndex);    T3.Charge = str2double(T3.Charge);
                s{shiftY,6} = sprintf('%0.2f', sum(T3.Charge)); %s{shiftY,8} = PriceTypeVar;
                
                % add summary to the summary sheet
                if obj.Settings.gui.GenerateSummaryFile
                    Summary{SummaryCounter,1} = s{shiftY,3}; 
                    Summary{SummaryCounter,2} = cell2mat(table2cell(T2(1, OrganizationIndex))); %s{shiftY,4}; 
                    Summary{SummaryCounter,3} = s{shiftY,4}; 
                    Summary{SummaryCounter,4} = s{shiftY,5}; 
                    Summary{SummaryCounter,5} = s{shiftY,6}; 
                    Summary{SummaryCounter,6} = PriceTypeVar; 
                    Summary{SummaryCounter,7} = s{8, 3};
                    SummaryCounter = SummaryCounter + 1;
                end
                
                xlswrite(fn, s, 'Invoice');
                %xlswrite2(fn, s, 'Invoice');
                
                writetable(T2, fn, 'Sheet', 'Full listing');
                 
                % remove default sheets 1,2,3
                objExcel = actxserver('Excel.Application');
                objExcel.Workbooks.Open(fullfile(fn)); % Full path is necessary!
                % Delete sheets 1, 2, 3.
                try
                    objExcel.ActiveWorkbook.Worksheets.Item(1).Delete;
                    objExcel.ActiveWorkbook.Worksheets.Item(1).Delete;
                    objExcel.ActiveWorkbook.Worksheets.Item(1).Delete;
                catch err
                    err;
                end
                sheet = objExcel.ActiveWorkbook.Sheets.Item(1);
                sheet.Activate;
                objExcel.ActiveSheet.Range('C1').Font.Bold = true;
                objExcel.ActiveSheet.Range('C1').Font.Size = 16;
                
                objExcel.ActiveSheet.Range('C13:F13').Font.Bold = true;
                objExcel.ActiveSheet.Range('C16:H16').Font.Bold = true;
                objExcel.ActiveSheet.Range('D3').Font.Bold = true;  % Billing address
                
                rangeText = sprintf('A1:A%d', shiftY);
                objExcel.ActiveSheet.Range(rangeText).Font.Bold = true;
                objExcel.ActiveSheet.Range(rangeText).HorizontalAlignment = 4;  % 2-left, 3-center, 4 - right
                
                objExcel.ActiveSheet.Columns.Item(1).columnWidth = 32; % 1st column width
                objExcel.ActiveSheet.Columns.Item(2).columnWidth = 1; % 2nd column width
                objExcel.ActiveSheet.Columns.Item(3).columnWidth = 35;
                objExcel.ActiveSheet.Columns.Item(4).columnWidth = 22;
                objExcel.ActiveSheet.Columns.Item(5).columnWidth = 18;
                objExcel.ActiveSheet.Columns.Item(6).columnWidth = 18;
                objExcel.ActiveSheet.Columns.Item(7).columnWidth = 18;
                objExcel.ActiveSheet.Columns.Item(8).columnWidth = 18;
                objExcel.ActiveSheet.Columns.Item(9).columnWidth = 18;
                
                % merge billing address cells
                objExcel.ActiveSheet.Range('E3:F11').MergeCells = 1;
                objExcel.ActiveSheet.Range('E3').VerticalAlignment = -4160; % align to top, see more https://docs.microsoft.com/en-us/office/vba/api/excel.xlvalign
                objExcel.ActiveSheet.Range('E3').WrapText = 1;
                objExcel.ActiveSheet.Range('C8:D9').MergeCells = 1;
                objExcel.ActiveSheet.Range('C8').VerticalAlignment = -4160;
                objExcel.ActiveSheet.Range('C8').WrapText = 1;
                
                % add borders
                objExcel.ActiveSheet.Range('A1:F1').Borders.Item('xlEdgeBottom').LineStyle = 1;
                for i=1:numel(lineVec)
                    rangeText = sprintf('A%d:F%d', lineVec(i), lineVec(i));
                    objExcel.ActiveSheet.Range(rangeText).Borders.Item('xlEdgeBottom').LineStyle = 1;
                    objExcel.ActiveSheet.Range(rangeText).Font.Bold = true;
                    rangeText = sprintf('A%d', lineVec(i));
                    objExcel.ActiveSheet.Range(rangeText).Font.Size = 12;
                end
                
                % highlight the summary section
                rangeText = sprintf('A%d:F%d', lineVec(end), lineVec(end));
                objExcel.ActiveSheet.Range(rangeText).Interior.ColorIndex = 19; %40; % RGB(r, g, b)
                rangeText = sprintf('F%d', lineVec(end)+1);
                objExcel.ActiveSheet.Range(rangeText).Interior.ColorIndex = 40; %45; % RGB(r, g, b)
                objExcel.ActiveSheet.Range(rangeText).Font.Bold = true;
                
                % add alignment
                rangeText = sprintf('D8:H%d', shiftY);
                objExcel.ActiveSheet.Range(rangeText).HorizontalAlignment = 3;  % 2-left, 3-center, 4 - right
                
                objExcel.PrintCommunication = 1;
                objExcel.ActiveSheet.PageSetup.Zoom = false;
                objExcel.ActiveSheet.PageSetup.FitToPagesWide = 1;
                %objExcel.ActiveSheet.PageSetup.PrintArea = sprintf('A1:F%d', size(s, 1));  % "$A$1:$C$5";
                
                % Save, close and clean up.
                objExcel.ActiveWorkbook.Save;
                objExcel.ActiveWorkbook.Close;
                objExcel.Quit;
                objExcel.delete;
                
%                 % save summary sheet
%                 % !!!!!! FOR DEBUG DELETE LATER !!!!!!!
%                 if obj.Settings.gui.GenerateSummaryFile
%                     fnTemplate = sprintf('%s_Summary.xlsx', dateString); %#ok<FNDSB>
%                     fnSummary = fullfile(obj.Settings.gui.OutputDirectory, fnTemplate); 
%                     if exist(fnSummary, 'file') == 2; delete(fnSummary); end
%                     xlswrite(fnSummary, Summary, 'InvoiceSummary');
%                 end
%                 % !!!!!! END OF DEBUG DELETE LATER !!!!!!!
                
                waitbar(splitId/numel(splitEntries), wb);
            end
            % save summary sheet
            if obj.Settings.gui.GenerateSummaryFile
                fnTemplate = sprintf('%s_Summary.xlsx', dateString); %#ok<FNDSB>
                fnSummary = fullfile(obj.Settings.gui.OutputDirectory, fnTemplate);
                waitbar(1, wb, sprintf('Writting the summary file\n%s', fnSummary));
                [~, sortedIndex] = sort(Summary(2:end,1));  % sort by group name
                Summary = [Summary(1,:); Summary(sortedIndex+1,:)];
                
                if exist(fnSummary, 'file') == 2; delete(fnSummary); end
                xlswrite(fnSummary, Summary, 'InvoiceSummary');
            end
            
            delete(wb);
            toc
        end
    end
    
end