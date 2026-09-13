Version =20
VersionRequired =20
Begin Form
    FilterOn = NotDefault
    DefaultView =0
    ViewsAllowed =1
    Filter ="ID = 1"
    RecSrcDt = Begin
        0x9f914306b67de640
    End
    RecordSource ="tCustomers"
    Caption ="Customer Detail"
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
        Begin Section
            Height =5400
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =200
                    Top =200
                    Width =1500
                    Height =315
                    Name ="lblID"
                    Caption ="ID"
                    LayoutCachedLeft =200
                    LayoutCachedTop =200
                    LayoutCachedWidth =1700
                    LayoutCachedHeight =515
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =1900
                    Top =200
                    Width =2000
                    Height =315
                    Name ="txtID"
                    ControlSource ="ID"

                    LayoutCachedLeft =1900
                    LayoutCachedTop =200
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =515
                End
                Begin Label
                    OverlapFlags =85
                    Left =200
                    Top =650
                    Width =1500
                    Height =315
                    Name ="lblFirstName"
                    Caption ="First Name"
                    LayoutCachedLeft =200
                    LayoutCachedTop =650
                    LayoutCachedWidth =1700
                    LayoutCachedHeight =965
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1900
                    Top =650
                    Width =3000
                    Height =315
                    TabIndex =1
                    Name ="txtFirstName"
                    ControlSource ="FirstName"

                    LayoutCachedLeft =1900
                    LayoutCachedTop =650
                    LayoutCachedWidth =4900
                    LayoutCachedHeight =965
                End
                Begin Label
                    OverlapFlags =85
                    Left =200
                    Top =1100
                    Width =1500
                    Height =315
                    Name ="lblLastName"
                    Caption ="Last Name"
                    LayoutCachedLeft =200
                    LayoutCachedTop =1100
                    LayoutCachedWidth =1700
                    LayoutCachedHeight =1415
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1900
                    Top =1100
                    Width =3000
                    Height =315
                    TabIndex =2
                    Name ="txtLastName"
                    ControlSource ="LastName"

                    LayoutCachedLeft =1900
                    LayoutCachedTop =1100
                    LayoutCachedWidth =4900
                    LayoutCachedHeight =1415
                End
                Begin Label
                    OverlapFlags =85
                    Left =200
                    Top =1550
                    Width =1500
                    Height =315
                    Name ="lblAddress"
                    Caption ="Address"
                    LayoutCachedLeft =200
                    LayoutCachedTop =1550
                    LayoutCachedWidth =1700
                    LayoutCachedHeight =1865
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1900
                    Top =1550
                    Width =5000
                    Height =900
                    TabIndex =3
                    Name ="txtAddress"
                    ControlSource ="Address"

                    LayoutCachedLeft =1900
                    LayoutCachedTop =1550
                    LayoutCachedWidth =6900
                    LayoutCachedHeight =2450
                End
                Begin Label
                    OverlapFlags =85
                    Left =200
                    Top =2600
                    Width =1500
                    Height =315
                    Name ="lblCity"
                    Caption ="City"
                    LayoutCachedLeft =200
                    LayoutCachedTop =2600
                    LayoutCachedWidth =1700
                    LayoutCachedHeight =2915
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1900
                    Top =2600
                    Width =3000
                    Height =315
                    TabIndex =4
                    Name ="txtCity"
                    ControlSource ="City"

                    LayoutCachedLeft =1900
                    LayoutCachedTop =2600
                    LayoutCachedWidth =4900
                    LayoutCachedHeight =2915
                End
                Begin Label
                    OverlapFlags =85
                    Left =200
                    Top =3050
                    Width =1500
                    Height =315
                    Name ="lblState"
                    Caption ="State"
                    LayoutCachedLeft =200
                    LayoutCachedTop =3050
                    LayoutCachedWidth =1700
                    LayoutCachedHeight =3365
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1900
                    Top =3050
                    Width =2000
                    Height =315
                    TabIndex =5
                    Name ="txtState"
                    ControlSource ="State"

                    LayoutCachedLeft =1900
                    LayoutCachedTop =3050
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =3365
                End
                Begin Label
                    OverlapFlags =85
                    Left =200
                    Top =3500
                    Width =1500
                    Height =315
                    Name ="lblPostCode"
                    Caption ="Post Code"
                    LayoutCachedLeft =200
                    LayoutCachedTop =3500
                    LayoutCachedWidth =1700
                    LayoutCachedHeight =3815
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1900
                    Top =3500
                    Width =2000
                    Height =315
                    TabIndex =6
                    Name ="txtPostCode"
                    ControlSource ="PostCode"

                    LayoutCachedLeft =1900
                    LayoutCachedTop =3500
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =3815
                End
                Begin Label
                    OverlapFlags =85
                    Left =200
                    Top =3950
                    Width =1500
                    Height =315
                    Name ="lblPhone"
                    Caption ="Phone"
                    LayoutCachedLeft =200
                    LayoutCachedTop =3950
                    LayoutCachedWidth =1700
                    LayoutCachedHeight =4265
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1900
                    Top =3950
                    Width =3000
                    Height =315
                    TabIndex =7
                    Name ="txtPhone"
                    ControlSource ="Phone"

                    LayoutCachedLeft =1900
                    LayoutCachedTop =3950
                    LayoutCachedWidth =4900
                    LayoutCachedHeight =4265
                End
                Begin Label
                    OverlapFlags =85
                    Left =200
                    Top =4400
                    Width =1500
                    Height =315
                    Name ="lblEmail"
                    Caption ="Email"
                    LayoutCachedLeft =200
                    LayoutCachedTop =4400
                    LayoutCachedWidth =1700
                    LayoutCachedHeight =4715
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1900
                    Top =4400
                    Width =5000
                    Height =315
                    TabIndex =8
                    Name ="txtEmail"
                    ControlSource ="Email"

                    LayoutCachedLeft =1900
                    LayoutCachedTop =4400
                    LayoutCachedWidth =6900
                    LayoutCachedHeight =4715
                End
            End
        End
        Begin FormFooter
            Height =600
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =200
                    Top =100
                    Width =2500
                    Height =440
                    Name ="cmdSaveClose"
                    Caption ="Save and Close"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =200
                    LayoutCachedTop =100
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =540
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2900
                    Top =100
                    Width =2500
                    Height =440
                    TabIndex =1
                    Name ="cmdDiscardClose"
                    Caption ="Discard and Close"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =2900
                    LayoutCachedTop =100
                    LayoutCachedWidth =5400
                    LayoutCachedHeight =540
                End
            End
        End
    End
End
CodeBehindForm
' See "fCustomerDetail.cls"
