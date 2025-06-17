﻿Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7389
    DatasheetFontHeight =11
    ItemSuffix =335
    Left =5550
    Top =3030
    Right =18705
    Bottom =14760
    RecSrcDt = Begin
        0x956642cd6e4ee640
    End
    Caption ="Install Add-in"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    OrderByOnLoad =0
    OrderByOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =60.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Line
            BorderLineStyle =0
            Width =1701
            BorderThemeColorIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            Width =1701
            Height =1701
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =7278
            Name ="Detailbereich"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2070
                    Top =3135
                    Width =4740
                    Height =300
                    TabIndex =5
                    Name ="txtAddInTitle"
                    FontName ="Tahoma"
                    GroupTable =1
                    RightPadding =567
                    BottomPadding =150
                    ShowDatePicker =0

                    LayoutCachedLeft =2070
                    LayoutCachedTop =3135
                    LayoutCachedWidth =6810
                    LayoutCachedHeight =3435
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =570
                            Top =3135
                            Width =1440
                            Height =300
                            Name ="lbltxtAddInName"
                            Caption ="Title"
                            FontName ="Tahoma"
                            GroupTable =1
                            LeftPadding =57
                            BottomPadding =150
                            HorizontalAnchor =2
                            LayoutCachedLeft =570
                            LayoutCachedTop =3135
                            LayoutCachedWidth =2010
                            LayoutCachedHeight =3435
                            RowStart =6
                            RowEnd =6
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2070
                    Top =570
                    Width =4740
                    Height =300
                    TabIndex =1
                    Name ="txtFileName"
                    FontName ="Tahoma"
                    GroupTable =1
                    RightPadding =567
                    BottomPadding =150
                    ShowDatePicker =0

                    LayoutCachedLeft =2070
                    LayoutCachedTop =570
                    LayoutCachedWidth =6810
                    LayoutCachedHeight =870
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =570
                            Top =570
                            Width =1440
                            Height =300
                            Name ="lblFileName"
                            Caption ="File name"
                            FontName ="Tahoma"
                            GroupTable =1
                            LeftPadding =57
                            BottomPadding =150
                            HorizontalAnchor =2
                            LayoutCachedLeft =570
                            LayoutCachedTop =570
                            LayoutCachedWidth =2010
                            LayoutCachedHeight =870
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2070
                    Top =3615
                    Width =4740
                    Height =300
                    TabIndex =6
                    Name ="txtAddInAuthor"
                    FontName ="Tahoma"
                    GroupTable =1
                    RightPadding =567
                    BottomPadding =150
                    ShowDatePicker =0

                    LayoutCachedLeft =2070
                    LayoutCachedTop =3615
                    LayoutCachedWidth =6810
                    LayoutCachedHeight =3915
                    RowStart =7
                    RowEnd =7
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =570
                            Top =3615
                            Width =1440
                            Height =300
                            Name ="lblAddInAuthor"
                            Caption ="Author"
                            FontName ="Tahoma"
                            GroupTable =1
                            LeftPadding =57
                            BottomPadding =150
                            HorizontalAnchor =2
                            LayoutCachedLeft =570
                            LayoutCachedTop =3615
                            LayoutCachedWidth =2010
                            LayoutCachedHeight =3915
                            RowStart =7
                            RowEnd =7
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2070
                    Top =4095
                    Width =4740
                    Height =300
                    TabIndex =7
                    Name ="txtAddInCompany"
                    FontName ="Tahoma"
                    GroupTable =1
                    RightPadding =567
                    BottomPadding =150
                    ShowDatePicker =0

                    LayoutCachedLeft =2070
                    LayoutCachedTop =4095
                    LayoutCachedWidth =6810
                    LayoutCachedHeight =4395
                    RowStart =8
                    RowEnd =8
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =570
                            Top =4095
                            Width =1440
                            Height =300
                            Name ="lblAddInCompany"
                            Caption ="Company"
                            FontName ="Tahoma"
                            GroupTable =1
                            LeftPadding =57
                            BottomPadding =150
                            HorizontalAnchor =2
                            LayoutCachedLeft =570
                            LayoutCachedTop =4095
                            LayoutCachedWidth =2010
                            LayoutCachedHeight =4395
                            RowStart =8
                            RowEnd =8
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2070
                    Top =4575
                    Width =4740
                    Height =1125
                    TabIndex =8
                    Name ="txtAddInComment"
                    FontName ="Tahoma"
                    GroupTable =1
                    RightPadding =567
                    BottomPadding =150
                    VerticalAnchor =2
                    ShowDatePicker =0

                    LayoutCachedLeft =2070
                    LayoutCachedTop =4575
                    LayoutCachedWidth =6810
                    LayoutCachedHeight =5700
                    RowStart =9
                    RowEnd =9
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =570
                            Top =4575
                            Width =1440
                            Height =1125
                            Name ="lblAddInComment"
                            Caption ="Comment"
                            FontName ="Tahoma"
                            GroupTable =1
                            LeftPadding =57
                            BottomPadding =150
                            HorizontalAnchor =2
                            VerticalAnchor =2
                            LayoutCachedLeft =570
                            LayoutCachedTop =4575
                            LayoutCachedWidth =2010
                            LayoutCachedHeight =5700
                            RowStart =9
                            RowEnd =9
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            GroupTable =1
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =570
                    Top =6240
                    Width =6240
                    Height =450
                    TabIndex =10
                    Name ="cmdInstallAddIn"
                    Caption ="Install Add-in"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    LeftPadding =57
                    RightPadding =567
                    BottomPadding =567
                    HorizontalAnchor =2

                    LayoutCachedLeft =570
                    LayoutCachedTop =6240
                    LayoutCachedWidth =6810
                    LayoutCachedHeight =6690
                    RowStart =11
                    RowEnd =11
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2070
                    Top =2340
                    Width =4740
                    Height =291
                    TabIndex =4
                    Name ="txtAddInStartFunction"
                    FontName ="Tahoma"
                    GroupTable =1
                    RightPadding =567
                    BottomPadding =150
                    ShowDatePicker =0

                    LayoutCachedLeft =2070
                    LayoutCachedTop =2340
                    LayoutCachedWidth =6810
                    LayoutCachedHeight =2631
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =570
                            Top =2340
                            Width =1440
                            Height =291
                            Name ="lblAddInStartFunction"
                            Caption ="Start Function"
                            FontName ="Tahoma"
                            GroupTable =1
                            LeftPadding =57
                            BottomPadding =150
                            HorizontalAnchor =2
                            LayoutCachedLeft =570
                            LayoutCachedTop =2340
                            LayoutCachedWidth =2010
                            LayoutCachedHeight =2631
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2070
                    Top =1860
                    Width =4740
                    Height =300
                    TabIndex =3
                    Name ="txtAddInRegPathName"
                    FontName ="Tahoma"
                    GroupTable =1
                    RightPadding =567
                    BottomPadding =150
                    ShowDatePicker =0

                    LayoutCachedLeft =2070
                    LayoutCachedTop =1860
                    LayoutCachedWidth =6810
                    LayoutCachedHeight =2160
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =570
                            Top =1860
                            Width =1440
                            Height =300
                            Name ="Bezeichnungsfeld105"
                            Caption ="Name"
                            FontName ="Tahoma"
                            GroupTable =1
                            LeftPadding =57
                            BottomPadding =150
                            HorizontalAnchor =2
                            LayoutCachedLeft =570
                            LayoutCachedTop =1860
                            LayoutCachedWidth =2010
                            LayoutCachedHeight =2160
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            GroupTable =1
                        End
                    End
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =570
                    Top =1530
                    Width =6240
                    Height =300
                    FontWeight =700
                    Name ="Bezeichnungsfeld112"
                    Caption ="USysRegInfo"
                    FontName ="Tahoma"
                    GroupTable =1
                    LeftPadding =57
                    RightPadding =567
                    BottomPadding =0
                    HorizontalAnchor =2
                    LayoutCachedLeft =570
                    LayoutCachedTop =1530
                    LayoutCachedWidth =6810
                    LayoutCachedHeight =1830
                    RowStart =2
                    RowEnd =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =570
                    Top =2805
                    Width =6240
                    Height =300
                    FontWeight =700
                    Name ="Bezeichnungsfeld150"
                    Caption ="Database properties"
                    FontName ="Tahoma"
                    GroupTable =1
                    LeftPadding =57
                    RightPadding =567
                    BottomPadding =0
                    HorizontalAnchor =2
                    LayoutCachedLeft =570
                    LayoutCachedTop =2805
                    LayoutCachedWidth =6810
                    LayoutCachedHeight =3105
                    RowStart =5
                    RowEnd =5
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2070
                    Top =1050
                    Width =4740
                    Height =300
                    TabIndex =2
                    Name ="txtAppTitle"
                    FontName ="Tahoma"
                    GroupTable =1
                    RightPadding =567
                    BottomPadding =150
                    ShowDatePicker =0

                    LayoutCachedLeft =2070
                    LayoutCachedTop =1050
                    LayoutCachedWidth =6810
                    LayoutCachedHeight =1350
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =570
                            Top =1050
                            Width =1440
                            Height =300
                            Name ="Label247"
                            Caption ="AppTitle"
                            FontName ="Tahoma"
                            GroupTable =1
                            LeftPadding =57
                            BottomPadding =150
                            HorizontalAnchor =2
                            LayoutCachedLeft =570
                            LayoutCachedTop =1050
                            LayoutCachedWidth =2010
                            LayoutCachedHeight =1350
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            GroupTable =1
                        End
                    End
                End
                Begin CommandButton
                    Transparent = NotDefault
                    OverlapFlags =85
                    Width =0
                    Height =0
                    Name ="sysFirst"
                    Caption ="-"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =3015
                    Top =5880
                    Width =3795
                    Height =300
                    TabIndex =9
                    Name ="cbCompileAddIn"
                    DefaultValue ="False"
                    GroupTable =1
                    RightPadding =567

                    LayoutCachedLeft =3015
                    LayoutCachedTop =5880
                    LayoutCachedWidth =6810
                    LayoutCachedHeight =6180
                    RowStart =10
                    RowEnd =10
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =570
                            Top =5880
                            Width =2415
                            Height =300
                            ForeColor =0
                            Name ="Label325"
                            Caption ="Install Add-in as accde"
                            GroupTable =1
                            LeftPadding =57
                            RightPadding =0
                            HorizontalAnchor =2
                            LayoutCachedLeft =570
                            LayoutCachedTop =5880
                            LayoutCachedWidth =2985
                            LayoutCachedHeight =6180
                            RowStart =10
                            RowEnd =10
                            ColumnEnd =1
                            LayoutGroup =1
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =114
                    Top =6916
                    Width =7139
                    Height =223
                    FontSize =8
                    Name ="lblVersionInfo"
                    HorizontalAnchor =2
                    LayoutCachedLeft =114
                    LayoutCachedTop =6916
                    LayoutCachedWidth =7253
                    LayoutCachedHeight =7139
                End
            End
        End
    End
End
CodeBehindForm
' See "InstallAddInForm.cls"
