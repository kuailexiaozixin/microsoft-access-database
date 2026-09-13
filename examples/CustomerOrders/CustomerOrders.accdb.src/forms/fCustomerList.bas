Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    ViewsAllowed =1
    Width =10000
    RecSrcDt = Begin
        0xfe554606b67de640
    End
    RecordSource ="qCustomers"
    Caption ="Customer List"
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
        Begin TextBox
            FELineBreak = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin FormHeader
            Height =400
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =100
                    Top =60
                    Width =600
                    Height =315
                    Name ="lblID"
                    Caption ="ID"
                    LayoutCachedLeft =100
                    LayoutCachedTop =60
                    LayoutCachedWidth =700
                    LayoutCachedHeight =375
                End
                Begin Label
                    OverlapFlags =85
                    Left =800
                    Top =60
                    Width =1800
                    Height =315
                    Name ="lblFirstName"
                    Caption ="First Name"
                    LayoutCachedLeft =800
                    LayoutCachedTop =60
                    LayoutCachedWidth =2600
                    LayoutCachedHeight =375
                End
                Begin Label
                    OverlapFlags =85
                    Left =2700
                    Top =60
                    Width =1800
                    Height =315
                    Name ="lblLastName"
                    Caption ="Last Name"
                    LayoutCachedLeft =2700
                    LayoutCachedTop =60
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =375
                End
                Begin Label
                    OverlapFlags =85
                    Left =4600
                    Top =60
                    Width =1500
                    Height =315
                    Name ="lblCity"
                    Caption ="City"
                    LayoutCachedLeft =4600
                    LayoutCachedTop =60
                    LayoutCachedWidth =6100
                    LayoutCachedHeight =375
                End
                Begin Label
                    OverlapFlags =85
                    Left =6200
                    Top =60
                    Width =1500
                    Height =315
                    Name ="lblPhone"
                    Caption ="Phone"
                    LayoutCachedLeft =6200
                    LayoutCachedTop =60
                    LayoutCachedWidth =7700
                    LayoutCachedHeight =375
                End
                Begin Label
                    OverlapFlags =85
                    Left =7800
                    Top =60
                    Width =2100
                    Height =315
                    Name ="lblEmail"
                    Caption ="Email"
                    LayoutCachedLeft =7800
                    LayoutCachedTop =60
                    LayoutCachedWidth =9900
                    LayoutCachedHeight =375
                End
            End
        End
        Begin Section
            Height =360
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =100
                    Width =600
                    Height =315
                    Name ="txtID"
                    ControlSource ="ID"

                    LayoutCachedLeft =100
                    LayoutCachedWidth =700
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =800
                    Width =1800
                    Height =315
                    TabIndex =1
                    Name ="txtFirstName"
                    ControlSource ="FirstName"

                    LayoutCachedLeft =800
                    LayoutCachedWidth =2600
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =2700
                    Width =1800
                    Height =315
                    TabIndex =2
                    Name ="txtLastName"
                    ControlSource ="LastName"

                    LayoutCachedLeft =2700
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =4600
                    Width =1500
                    Height =315
                    TabIndex =3
                    Name ="txtCity"
                    ControlSource ="City"

                    LayoutCachedLeft =4600
                    LayoutCachedWidth =6100
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =6200
                    Width =1500
                    Height =315
                    TabIndex =4
                    Name ="txtPhone"
                    ControlSource ="Phone"

                    LayoutCachedLeft =6200
                    LayoutCachedWidth =7700
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =7800
                    Width =2100
                    Height =315
                    TabIndex =5
                    Name ="txtEmail"
                    ControlSource ="Email"

                    LayoutCachedLeft =7800
                    LayoutCachedWidth =9900
                    LayoutCachedHeight =315
                End
            End
        End
        Begin FormFooter
            Height =600
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =100
                    Top =100
                    Width =2000
                    Height =440
                    Name ="cmdOpen"
                    Caption ="Open"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =100
                    LayoutCachedTop =100
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =540
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2300
                    Top =100
                    Width =2000
                    Height =440
                    TabIndex =1
                    Name ="cmdAdd"
                    Caption ="Add New"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =2300
                    LayoutCachedTop =100
                    LayoutCachedWidth =4300
                    LayoutCachedHeight =540
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4500
                    Top =100
                    Width =2000
                    Height =440
                    TabIndex =2
                    Name ="cmdClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4500
                    LayoutCachedTop =100
                    LayoutCachedWidth =6500
                    LayoutCachedHeight =540
                End
            End
        End
    End
End
CodeBehindForm
' See "fCustomerList.cls"
