{ ****************************************************************************
  Kill Outlook Tasks

  After migrating my wife to a new computer system, I noticed that Outlook
  reported that it had 349,000 events. Deleting all of them by hand would
  be a nightmare, so I decided to do it with this code.

  Learned how to work with the Outlook calendar from Delphi using:
  http://www.scalabium.com/faq/dct0128.htm

  Learned how to access Outlook items using:
  http://www.scalabium.com/faq/dct0121.htm

  John M. Wargo
  October 25, 2018
  **************************************************************************** }
unit kill_task;

interface

uses
  ComObj,
  System.SysUtils, System.Variants, System.Classes,
  Winapi.Windows, Winapi.Messages,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ComCtrls,
  Vcl.StdCtrls;

type
  TfrmMain = class(TForm)
    StatusBar1: TStatusBar;
    output: TMemo;
    procedure FormActivate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

{$R *.dfm}

procedure TfrmMain.FormActivate(Sender: TObject);
const
  // the task type we're looking for in Outlook
  olTaskItem = $00000003;
var
  outlook, oiItem, ns, folder: OLEVariant;
  taskStatus: Boolean;
  i, numItems: Integer;
  msgText: String;

  { to find a default Calendar folder }
  function GetCalendarFolder(folder: OLEVariant): OLEVariant;
  var
    i: Integer;
  begin
    for i := 1 to folder.Count do
    begin
      if (folder.Item[i].DefaultItemType = olTaskItem) then
        result := folder.Item[i]
      else
        result := GetCalendarFolder(folder.Item[i].Folders);
      if not VarIsNull(result) and not VarIsEmpty(result) then
        break;
    end;
  end;

begin
  output.Lines.add('Creating Outlook Object');
  // initialize a connection to Outlook
  outlook := CreateOLEObject('Outlook.Application');
  // get the MAPI namespace
  ns := outlook.GetNamespace('MAPI');
  // get a default Contacts folder
  folder := GetCalendarFolder(ns.Folders);
  output.Lines.add('Default calendar folder: ' + string(folder));
  // if  Calendar folder is found
  if VarIsNull(folder) and not VarIsEmpty(folder) then
  begin
    // Then tell the user
    msgText := 'Unable to determine default calendar folder';
    ShowMessage(msgText);
    output.Lines.add(msgText);
  end
  else
  begin
    // Process entries in the folder
    output.Lines.add(Format('Searching "%s" folder', [folder]));
    numItems := folder.Items.Count;
    if (numItems > 0) then
    begin
      output.Lines.add(Format('Found %d items', [numItems]));
      // Have to go backwards through the list to delete them apparently
      for i := numItems downto 1 do
      begin
        oiItem := folder.Items[i];
        taskStatus := oiItem.complete;
        // Is the task completed?
        if taskStatus then
        begin
          msgText := Format('Deleting #%d: %s', [i, oiItem.subject]);
          output.Lines.add(msgText);
          // Uncomment the following line when you're ready to delete entries
          // oiItem.Delete;
          //Let the app process messages before kicking off the next one
          Application.ProcessMessages();
        end;
      end;
      output.Lines.add('All done!');
    end
    else
    begin
      output.Lines.add('No entries found');
    end;
  end;

end;

end.
