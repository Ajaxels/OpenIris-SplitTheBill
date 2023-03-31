function SplitTheBill()
% @mainpage SplitTheBill
% @section intro Introduction
% @b SortBills is a Matlab based program for sorting the bills that are
% produced with the OpenIris reservation system (https://openiris.io).
% @section features Key Features
% - loading excel spreadsheets 
% - sorting
% - export to multiple sheets, bills
% @section description Description
% Automatization of billing

% Copyright (C) 2019-2020 Ilya Belevich, University of Helsinki (ilya.belevich @ helsinki.fi)
% The MIT License (https://opensource.org/licenses/MIT)

% Updates:
% 2022.1: fixed extraction of products that were added to the booking slots

% turn off warnings
warning('off', 'MATLAB:ui:javaframe:PropertyToBeRemoved'); 

if ~isdeployed
    func_name='SplitTheBill.m';
    func_dir=which(func_name);
    func_dir=fileparts(func_dir);
    addpath(func_dir);
    addpath(fullfile(func_dir, 'Tools'));
    % addpath(fullfile(func_dir, 'Classes'));
end
version = 'ver. 2023.01 (31.03.2023)';
if isdeployed; version = [version ' Academic version']; end
model = Model();     % initialize the model
controller = Controller(model, version);  % initialize controller
