Version =20
VersionRequired =20
Begin Form
    ViewsAllowed =1
    Width =10000
    RecSrcDt = Begin
        0xc0304b06b67de640
    End
    RecordSource ="tOrderItems"
    Caption ="Order Items"
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
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
                    Width =4000
                    Height =315
                    Name ="lblItemDescription"
                    Caption ="Item Description"
                    LayoutCachedLeft =100
                    LayoutCachedTop =60
                    LayoutCachedWidth =4100
                    LayoutCachedHeight =375
                End
                Begin Label
                    OverlapFlags =85
                    Left =4200
                    Top =60
                    Width =2000
                    Height =315
                    Name ="lblUnitPriceNet"
                    Caption ="Unit Price (Net)"
                    LayoutCachedLeft =4200
                    LayoutCachedTop =60
                    LayoutCachedWidth =6200
                    LayoutCachedHeight =375
                End
                Begin Label
                    OverlapFlags =85
                    Left =6300
                    Top =60
                    Width =1500
                    Height =315
                    Name ="lblQuantity"
                    Caption ="Quantity"
                    LayoutCachedLeft =6300
                    LayoutCachedTop =60
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =375
                End
                Begin Label
                    OverlapFlags =85
                    Left =7900
                    Top =60
                    Width =2000
                    Height =315
                    Name ="lblSubTotalNet"
                    Caption ="Sub Total (Net)"
                    LayoutCachedLeft =7900
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
                    OverlapFlags =85
                    Left =100
                    Width =4000
                    Height =315
                    Name ="txtItemDescription"
                    ControlSource ="ItemDescription"

                    LayoutCachedLeft =100
                    LayoutCachedWidth =4100
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =4200
                    Width =2000
                    Height =315
                    TabIndex =1
                    Name ="txtUnitPriceNet"
                    ControlSource ="UnitPriceNet"
                    Format ="$#,##0.00;-$#,##0.00"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =4200
                    LayoutCachedWidth =6200
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =6300
                    Width =1500
                    Height =315
                    TabIndex =2
                    Name ="txtQuantity"
                    ControlSource ="Quantity"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =6300
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =7900
                    Width =2000
                    Height =315
                    TabIndex =3
                    Name ="txtSubTotalNet"
                    ControlSource ="SubTotalNet"
                    Format ="$#,##0.00;-$#,##0.00"

                    LayoutCachedLeft =7900
                    LayoutCachedWidth =9900
                    LayoutCachedHeight =315
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "fOrderItemsSub.cls"
