{------------------------------------------------------------------------

     Smart Browse For Folder

     example of smart usage of BrowseForFolder API function
     including
     - repositioning and resizing browse window
     - adding a listbox, static elements and a button
     - catching button click
     - filling the listbox with file names
     - custom information field
     - custom condition for allowing folder selection
     - creating new folder
     - !!! REFRESHING TREE !!! after folder creation
       (thanks to Leonid Kunin for his idea published at
        http://codeguru.earthweb.com)

     Copyright (C) Konstantin Polyakov, 2001
     
     FIDO:   2:5030/542.251
     e-mail: kpolyakov@mail.ru
     Web:    http://kpolyakov.newmail.ru

------------------------------------------------------------------------}
program SmartBff;

uses Windows,  Messages,  ActiveX,  ShlObj,  CommCtrl,  Dialogs;

type AWndProc = function (Wnd: HWND; Msg: UINT; wParam: WPARAM; lParam: LPARAM):
                       LRESULT; stdcall;

const
    ID_CREATEBTN = 100;
    FileMask = '*.*';

var MainWnd, TreeWnd, LBoxWnd, StatusWnd,
    DirLabel, CreateBtn: HWND;
    OldWndProc: AWndProc;
    PathSelected: string;

//-------------------------------------------------------------------
//       FILL LISTBOX
//-------------------------------------------------------------------
procedure FillListBox(LBoxWnd: HWND; Path, Mask: string);
var FindHandle: THandle;
    FindData: TWin32FindData;
begin
    SendMessage(LBoxWnd, LB_RESETCONTENT, 0, 0);
    if Path = '' then Exit;
    Path := Path + Mask;
    FindHandle := FindFirstFile(PChar(Path), FindData);
    while FindHandle <> INVALID_HANDLE_VALUE do begin
       if (FILE_ATTRIBUTE_DIRECTORY and FindData.dwFileAttributes) = 0 then begin
         SendMessage(LBoxWnd, LB_ADDSTRING, 0, Longint(@FindData.cFileName[0]));
       end;
       if not FindNextFile(FindHandle, FindData) then begin
         Windows.FindClose(FindHandle);
         break;
       end;
    end;
end;

//-------------------------------------------------------------------
//       GET STATUS TEXT
//-------------------------------------------------------------------
function GetStatusText(var Enable: integer; Path: string): string;
begin
    Result := '';
    if Enable = 0 then begin
       Result := 'Можно выбирать каталоги только на жестких дисках';
       EnableWindow(CreateBtn, False);
       Exit;
    end;
    EnableWindow(CreateBtn, True);
    if SendMessage(LBoxWnd, LB_GETCOUNT, 0, 0) = 0 then begin
       Enable := 0;
       Result := 'В этой папке нет нужных файлов.';
    end;
end;

//-------------------------------------------------------------------
//       BROWSE WND PROC
//-------------------------------------------------------------------
function BrowseWndProc(Wnd: HWND; Msg: UINT; wParam: WPARAM; lParam: LPARAM):
                       LRESULT; stdcall;
var wNotifyCode: integer;
    wID: integer;
    CurItem: HTreeItem;
    Item: TTVItem;
    Folder, FullPath: string;
begin
   if Msg = WM_COMMAND then begin
      wNotifyCode := HIWORD(wParam);
      wID := LOWORD(wParam);
      if (wNotifyCode = BN_CLICKED) and (wID = ID_CREATEBTN) then begin
         Folder := 'New Folder';
         if InputQuery('New folder', 'Enter new folder name', Folder) then begin
           FullPath := PathSelected + '\' + Folder;
           CreateDirectory(PChar(FullPath), nil);
           CurItem := TreeView_GetSelection(TreeWnd);
           TreeView_Expand(TreeWnd, CurItem, TVE_COLLAPSE or TVE_COLLAPSERESET);
           ZeroMemory(@Item, sizeof(Item));
           Item.hItem := CurItem;
           Item.mask := TVIF_HANDLE or TVIF_CHILDREN;
           Item.cChildren := I_CHILDRENCALLBACK;
           TreeView_SetItem(TreeWnd, Item);
           SendMessage(MainWnd, BFFM_SETSELECTION, 1, integer(PChar(FullPath)) );
           Windows.SetFocus(TreeWnd);
         end;
         Result := 0;
         Exit;
      end;
   end;
   Result := OldWndProc(Wnd, Msg, wParam, lParam);
end;

//-------------------------------------------------------------------
//       MOVE CHILD UP
//-------------------------------------------------------------------
function MoveChildUp(CWnd: HWND; shiftY: integer): longbool; stdcall;
var rct: TRect;
begin
  if CWnd <> DirLabel then begin
     GetWindowRect(CWnd, rct);
     ScreenToClient(MainWnd, rct.TopLeft);
     SetWindowPos(CWnd, 0, rct.Left, rct.Top - shiftY, 0, 0,
                        SWP_NOOWNERZORDER or SWP_NOSIZE );
  end;
  Result := True;
end;

//-------------------------------------------------------------------
//       CREATE BROWSE WINDOW
//-------------------------------------------------------------------
procedure CreateBrowseWindow(Wnd: HWND);
const topMargin = 20;
      wLBox = 125;
var rct, rctStatic, rctTree, rctLBox, rctBtn: TRect;
    hLBox, lLBox, tLBox: integer;
    w, h, wBtn, dh: integer;
    FileStatic, BtnWnd: HWND;
    FontHandle: THandle;
    Style: Integer;
begin
  MainWnd := Wnd;
  GetWindowRect(Wnd, rct);
  w := rct.Right - rct.Left + 135;
  h := rct.Bottom - rct.Top;

             // find treeview
  TreeWnd := FindWindowEx(Wnd, 0, PChar('SysTreeView32'), nil);
  GetWindowRect(TreeWnd, rctTree);
  ScreenToClient(Wnd, rctTree.TopLeft);
  ScreenToClient(Wnd, rctTree.BottomRight);
             // HideSelection := False
  Style := GetWindowLong(TreeWnd, GWL_STYLE);
  SetWindowLong(TreeWnd, GWL_STYLE, Style or TVS_SHOWSELALWAYS );

            // store treeview font handle
  FontHandle := SendMessage(TreeWnd, WM_GETFONT, 0, 0);

            // find static text element
  DirLabel := FindWindowEx(Wnd, 0, PChar('Static'), nil);
  GetWindowRect(DirLabel, rctStatic);
  ScreenToClient(Wnd, rctStatic.TopLeft);
  dh := rctTree.Top - rctStatic.Top - topMargin;

             // find button
  BtnWnd := FindWindowEx(Wnd, 0, PChar('Button'), nil);
  GetWindowRect(BtnWnd, rctBtn);
  ScreenToClient(Wnd, rctBtn.TopLeft);
  ScreenToClient(Wnd, rctBtn.BottomRight);

            // move all child windows up by 'dh'
  EnumChildWindows(Wnd, @MoveChildUp, dh);

            // resize static text
  SetWindowPos(DirLabel, 0, 0, 0,
               rctTree.Right - rctTree.Left,
               rctTree.Top - dh - rctStatic.Top,
               SWP_NOOWNERZORDER or SWP_NOMOVE );

           //  create listbox
  hLBox := rctTree.Bottom - rctTree.Top - 50;
  lLBox := rctTree.Right + 10;
  tLBox := rctStatic.Top + topMargin;
  LBoxWnd := CreateWindow('listbox', nil,
                    WS_VISIBLE or WS_CHILD or
                    LBS_STANDARD or LBS_NOINTEGRALHEIGHT or LBS_NOSEL,
                    lLBox, tLBox, wLBox, hLBox,
                    Wnd, 0, hInstance, nil);
  SendMessage(LBoxWnd, WM_SETFONT, FontHandle, 1);

           //  create additional static element
  FileStatic := CreateWindow('static', 'Файлы с данными',
                    WS_VISIBLE or SS_SIMPLE or WS_CHILD,
                    rctTree.Right+10, rctStatic.Top,
                    125,
                    rctTree.Top - rctStatic.Top,
                    Wnd,
                    0, hInstance, nil);
  SendMessage(FileStatic, WM_SETFONT, FontHandle, 1);

               // new comment static
  StatusWnd := CreateWindow('static', '',
                    WS_VISIBLE Or SS_LEFT Or WS_CHILD,
                    lLBox, tLBox + hLBox + 10,
                    wLBox, 40,
                    Wnd,
                    0, hInstance, nil);
  SendMessage(StatusWnd, WM_SETFONT, FontHandle, 1);

              // new button 'Create'
  GetWindowRect(LBoxWnd, rctLBox);
  ScreenToClient(Wnd, rctLBox.BottomRight);
  wBtn := rctBtn.Right - rctBtn.Left;
  CreateBtn := CreateWindow('button', 'Create',
                WS_VISIBLE Or WS_CHILD,
                rctLBox.Right - wBtn, rctBtn.Top - dh,
                wBtn,
                rctBtn.Bottom - rctBtn.Top,
                Wnd,
                ID_CREATEBTN, hInstance, nil);
  SendMessage(CreateBtn, WM_SETFONT, FontHandle, 1);

              // replace window procedure
  OldWndProc := AWndProc(GetWindowLong(Wnd, GWL_WNDPROC));
  SetWindowLong(Wnd, GWL_WNDPROC, Longint(@BrowseWndProc));

           // place window at the screen center
  SetWindowPos(Wnd, HWND_TOP, (800 - w) div 2,
               (600 - h) div 2 - 20, w, h - dh, 0 );
           // change window title
  SetWindowText(Wnd, PChar('Выбор каталога с данными'));
end;

//-------------------------------------------------------------------
//       BROWSE CALLBACK PROC
//-------------------------------------------------------------------
function BrowseCallbackProc( Wnd : THandle; uMsg : UINT;
                             lParam : Integer; lpData : Pointer ) : Integer; stdcall;
var Path :   array[0..MAX_PATH-1] of Char;
    RootDir, StatusText: string;
    Enable:  Integer;
begin
 case uMsg of
   BFFM_INITIALIZED: begin
        CreateBrowseWindow(Wnd);
        if Assigned( lpData ) then
           SendMessage( Wnd, BFFM_SETSELECTION, 1, Integer(lpData) );
      end;
   BFFM_SELCHANGED: begin
        Enable := 0;
        StatusText := '';
        if Assigned( lpData ) then begin
           SHGetPathFromIDList( PItemIdList(lparam), @Path[0] );
           SetWindowText(DirLabel, Path);
           PathSelected := Path;
           RootDir := Copy( Path, 1, Pos('\', Path ) );
           if GetDriveType( PChar( RootDir ) ) = DRIVE_FIXED then
              Enable := 1;
           FillListBox(LBoxWnd, Path, '\' + FileMask);
           StatusText := GetStatusText(Enable, Path);
           SetWindowText(StatusWnd, PChar(StatusText));
        end;
        SendMessage( Wnd, BFFM_ENABLEOK, 0, Enable );
      end;
 end;
 Result := 0;
end;

//-------------------------------------------------------------------
//       CHOOSE FOLDER
//-------------------------------------------------------------------
function ChooseFolder(Title, StartPath: string; Flags: UINT): string;
var bi: TBrowseInfo;
    buf: PChar;
    DrivesPIDL: PItemIDList;
    ItemIDList: PItemIDList;
    ShellMalloc: IMalloc;
begin
    Result:='';

//    if not DirectoryExists(StartPath) then
//       StartPath := ExtractFileDir(ParamStr(0));

    If (ShGetMalloc(ShellMalloc) <> S_OK) or (ShellMalloc = nil) then Exit;

    SHGetSpecialFolderLocation( 0, CSIDL_DRIVES, DrivesPIDL );
    buf:= ShellMalloc.Alloc(MAX_PATH);

    FillChar(bi, sizeof(bi), 0);
    try
      bi.hwndOwner := 0; //Application.Handle;
      bi.pidlRoot := DrivesPIDL;
      bi.pszDisplayName := @buf[1];
      bi.lpszTitle := PChar(title);
      bi.ulFlags := BIF_RETURNONLYFSDIRS or Flags;
      bi.lpfn := @BrowseCallbackProc;
      bi.iImage := 0;
      bi.lParam := Integer( StartPath );
      ItemIDList := ShBrowseForFolder(bi);
      If ItemIDList <> nil then begin
        ShGetPathFromIDList(ItemIDList, buf);
        ShellMalloc.Free(ItemIDList);
        Result:=buf;
      end;
    finally
      ShellMalloc.Free(buf);
    end;

end;

//-------------------------------------------------------------------
//       MAIN PROGRAM
//-------------------------------------------------------------------
begin
   ChooseFolder('Выбор каталога с файлами *.DAT', 'C:\', 0 );
end.
