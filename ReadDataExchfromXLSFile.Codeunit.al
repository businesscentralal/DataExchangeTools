codeunit 50100 "Read Data Exch. from XLS File"
{
    TableNo = "Data Exch.";

    trigger OnRun()
    var
        TempBlob: Codeunit "Temp Blob";
        FileMgt: Codeunit "File Management";
        RecordRef: RecordRef;
    begin
        OnBeforeFileImport(TempBlob, "File Name");

        if not TempBlob.HasValue then
            "File Name" := CopyStr(
                FileMgt.BLOBImportWithFilter(TempBlob, ImportBankStmtTxt, '', FileFilterTxt, FileFilterExtensionTxt), 1, 250);

        if "File Name" <> '' then begin
            RecordRef.GetTable(Rec);
            TempBlob.ToRecordRef(RecordRef, FieldNo("File Content"));
            RecordRef.SetTable(Rec);
        end;
    end;

    var
        ImportBankStmtTxt: Label 'Select a file to import';
        FileFilterTxt: Label 'All Files(*.*)|*.*|Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls';
        FileFilterExtensionTxt: Label 'xlsx,xls', Locked = true;

    [IntegrationEvent(false, false)]
    local procedure OnBeforeFileImport(var TempBlob: Codeunit "Temp Blob"; var FileName: Text)
    begin
    end;
}

