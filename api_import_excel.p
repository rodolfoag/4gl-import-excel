/**
 * Author: Rodolfo Goncalves (github.com/rodolfoag/4gl-import-excel)
 * Program API to extract data from Excel files trough dynamic temp-table definitions
 * Obs.: in order to use this api via Webspeed or AppServer, make sure that SYSTEM can uses Excel
 *       interfaces by applying the following configuration:
 *       - on Windows 2008 Server x64: make this folder C:\Windows\SysWOW64\config\systemprofile\Desktop
 *       - on Windows 2008 Server x86: make this folder C:\Windows\System32\config\systemprofile\Desktop
 */

/* Local vars */
def var ch-excel       as com-handle no-undo.
def var ch-sheet       as com-handle no-undo.
def var i-current-line as int        no-undo init 1.
def var i-current-col  as int        no-undo.

/* Functions */
function fi-cell-value returns char (p-field as handle):
    case p-field:data-type:
        when "character" then
            return string( string( ch-sheet:cells(i-current-line, i-current-col):value ), p-field:format ).
        otherwise do:
            return ch-sheet:cells(i-current-line, i-current-col):value.
        end.
    end case.
end.

/* Procedures */
procedure pi-import-excel-file:
    def input parameter p-file-path     as char   no-undo.
    def input parameter p-ignore-header as log    no-undo.
    def input-output parameter table-handle p-tt-handle.

    def var h-buffer     as handle     no-undo.
    def var i-num-fields as int        no-undo.
    def var h-field      as handle     no-undo.
    def var i-col        as int        no-undo.

    assign h-buffer     = p-tt-handle:default-buffer-handle.
           i-num-fields = h-buffer:num-fields.

    if p-ignore-header then
        assign i-current-line = 2.

    /* Open File - invisible, no alerts */
    create "Excel.Application" ch-excel. 
    ch-excel:visible        = no.
    ch-excel:DisplayAlerts  = no.
    ch-excel:ScreenUpdating = no.
    ch-excel:Workbooks:open(p-file-path).
    ch-sheet = ch-excel:Sheets:item(1).

    repeat: /* lines */
        /* discard when the first line cell is empty */
        if ch-sheet:cells(i-current-line, 1):value = ? then
            leave.

        h-buffer:buffer-create().

        assign i-current-col = 1.

        do i-col = 1 to i-num-fields:
            assign h-field = h-buffer:buffer-field(i-col).

            if h-field:extent = 0 then
                assign h-field:buffer-value = fi-cell-value(h-field)
                       i-current-col        = i-current-col + 1.
            else
                run pi-handle-array(input h-field).
                                    
        end.

        assign i-current-line = i-current-line + 1.
    end.

    /* Quit - enable alerts */
    ch-excel:DisplayAlerts = yes.
    ch-excel:quit().
    release object ch-sheet.
    release object ch-excel.
    assign ch-sheet = ?
           ch-excel = ?.
end.

procedure pi-handle-array:
    def input parameter p-field  as handle     no-undo.

    def var i-col   as int  no-undo.
    def var c-value as char no-undo extent.

    assign extent(c-value) = p-field:extent.

    do i-col = 1 to p-field:extent:
        assign c-value[i-col] = fi-cell-value(p-field)
               i-current-col  = i-current-col + 1.
    end.

    assign p-field:buffer-value = c-value.
end.
