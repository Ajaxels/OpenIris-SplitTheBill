function primaryStruct = mibConcatenateStructures(primaryStruct, secondaryStruct)
% function primaryStruct = mibConcatenateStructures(primaryStruct, secondaryStruct)
% update fields of  primaryStruct using the fields of secondaryStruct
%
% Parameters:
% primaryStruct: primary structure that should be updates
% secondaryStruct: secondary structure that should be concatenated into the primary structure
%
% Return value:
% primaryStruct: updated primary structure

%| 
% @b Examples:
% @code preferences = mibConcatenateStructures(preferences, newPreferences); //updates fields of obj.mibModel.preferences with fields from mib_pars.preferences @endcode
%
% Copyright (C) 2020 Ilya Belevich, University of Helsinki (ilya.belevich @ helsinki.fi)
% The MIT License (https://opensource.org/licenses/MIT)
% 
% Updates
% 

secFieldsList = fieldnames(secondaryStruct);

for fieldId = 1:length(secFieldsList)
    if isstruct(secondaryStruct.(secFieldsList{fieldId}))
        if isfield(primaryStruct, secFieldsList{fieldId})
            primaryStruct.(secFieldsList{fieldId}) = ...
                mibConcatenateStructures(primaryStruct.(secFieldsList{fieldId}), secondaryStruct.(secFieldsList{fieldId}));
        else
            primaryStruct.(secFieldsList{fieldId}) = secondaryStruct.(secFieldsList{fieldId});
        end
    else
        primaryStruct.(secFieldsList{fieldId}) = secondaryStruct.(secFieldsList{fieldId});
    end
end
end