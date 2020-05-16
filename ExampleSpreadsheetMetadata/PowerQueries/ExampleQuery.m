let
    Source = {1..10},
    ConvertToTable = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    RenamedCols = Table.RenameColumns(ConvertToTable,{{"Column1", "MyCol"}}),
    ChangedType = Table.TransformColumnTypes(RenamedCols,{{"MyCol", type number}}),
    AddCol = Table.AddColumn(ChangedType, "Power", each Number.Power([MyCol], 3), type number)
in
    AddCol