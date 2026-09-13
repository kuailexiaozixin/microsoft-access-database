Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =1
    Width =5000
    RecSrcDt = Begin
        0x3d384f06b67de640
    End
    Caption ="Customer Orders"
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin Section
            Height =2400
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =1500
                    Top =500
                    Width =2000
                    Height =440
                    Name ="cmdCustomers"
                    Caption ="Customers"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =1500
                    LayoutCachedTop =500
                    LayoutCachedWidth =3500
                    LayoutCachedHeight =940
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1500
                    Top =1100
                    Width =2000
                    Height =440
                    TabIndex =1
                    Name ="cmdOrders"
                    Caption ="Orders"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =1500
                    LayoutCachedTop =1100
                    LayoutCachedWidth =3500
                    LayoutCachedHeight =1540
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1500
                    Top =1700
                    Width =2000
                    Height =440
                    TabIndex =2
                    Name ="cmdExit"
                    Caption ="Exit"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =1500
                    LayoutCachedTop =1700
                    LayoutCachedWidth =3500
                    LayoutCachedHeight =2140
                End
            End
        End
    End
End
CodeBehindForm
' See "fStartup.cls"
