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
version = 'ver. 2020.01 (17.12.2020)';
model = Model();     % initialize the model
controller = Controller(model, version);  % initialize controller
