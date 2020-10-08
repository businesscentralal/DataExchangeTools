codeunit 50101 "Read Fields From XLS File"
{
    Permissions = tabledata "Data Exch. Field" = i;
    TableNo = "Data Exch.";

    trigger OnRun()
    var
        DataExchDef: Record "Data Exch. Def";
        ExcelBuffer: Record "Excel Buffer" temporary;
        InStr: InStream;
    begin
        DataExchDef.Get(Rec."Data Exch. Def Code");
        DataExchDef.TestField("Header Tag");
        Rec."File Content".CreateInStream(InStr);
        ExcelBuffer.OpenBookStream(InStr, DataExchDef."Header Tag");
        ExcelBuffer.ReadSheet();
        ReadExcelBufferToDataExchFields(Rec, ExcelBuffer, DataExchDef."Header Lines" + 1);
    end;

    local procedure ReadExcelBufferToDataExchFields(DataExch: Record "Data Exch."; var ExcelBuffer: Record "Excel Buffer"; StartLineNo: integer)
    var
        DataExchField: Record "Data Exch. Field";
        DataExchColDef: Record "Data Exch. Column Def";
        LineNo: Integer;
    begin
        DataExchColDef.SetRange("Data Exch. Def Code", DataExch."Data Exch. Def Code");
        DataExchColDef.SetRange("Data Exch. Line Def Code", DataExch."Data Exch. Line Def Code");
        for LineNo := StartLineNo to GetNoOfRows(ExcelBuffer) do begin
            DataExchColDef.FindSet();
            repeat
                if ExcelBuffer.Get(LineNo, DataExchColDef."Column No.") then
                    DataExchField.InsertRec(DataExch."Entry No.", LineNo, DataExchColDef."Column No.", ExcelBuffer."Cell Value as Text", DataExch."Data Exch. Line Def Code");
            until DataExchColDef.Next() = 0;
        end;


    end;

    local procedure GetNoOfRows(var ExcelBuffer: Record "Excel Buffer"): Integer
    begin
        ExcelBuffer.SetCurrentKey("Row No.");
        if ExcelBuffer.FindLast() then
            exit(ExcelBuffer."Row No.");
    end;
}