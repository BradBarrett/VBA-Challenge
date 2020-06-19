Sub Stocktickercalc()

dim stockvolume as double
dim stockticker as string
dim stockopen as double
dim stocksummaryrow as integer
dim lastrow as long
dim ws as Worksheet
dim rownum as long



For Each ws In Worksheets
'stock percent change
stockperchg = 0
'else counter
elsecount = 0
'process counter
processcount = 0
'set volume to zero
stockvolume = 0
'set summary
stocksummaryrow = 2
' set Yearly Change variable
stockchange = 0
'set last row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ws.cells(1,11).value = "Ticker"
    ws.cells(1,12).value = "Yearly Change"
    ws.cells(1,13).value = "Percent Change"
    ws.cells(1,14).value = "Total Stock Volume"
    
    For i = 2 to lastrow
    
        'If look foward check for not equals
        if ws.cells(i+1,1).value <> ws.cells(i,1).value then
            'Process count add
            
            
            'calculate Stock change
            stockchange = ws.cells(i,6).value - ws.cells((i-(elsecount)),3)
            
                'if divzero error
                if ws.cells((i-(elsecount)),3).value = 0 then

                ws.range("M" & stocksummaryrow).value = "Not Active"

                else
            
                'calculate stock % change
                stockperchg = stockchange / ws.cells((i-(elsecount)),3)
                'print stock % change
                ws.range("M" & stocksummaryrow).value = FormatPercent(stockperchg)

                end if

            'add final volume
            stockvolume = stockvolume + ws.cells(i,7).value

            'set stock ticker
            stockticker = ws.cells(i,1).value        
        
            'print stock ticker
            ws.range("K" & stocksummaryrow).value = stockticker

            'print stock volume
            ws.range("N" & stocksummaryrow).value = stockvolume

            'Format Stock Volume
            ws.range("N" & stocksummaryrow).NumberFormat = "#,##0"

            'print stock change
            ws.range("L" & stocksummaryrow).value = stockchange
            'format colors
                if stockchange < 0 then

                ws.range("L" & stocksummaryrow).Interior.ColorIndex = 3

                else

                ws.range("L" & stocksummaryrow).Interior.ColorIndex = 4

                end if 

            'add row
            stocksummaryrow = stocksummaryrow + 1 
        
            'clear stock volume
            stockvolume = 0

            'reset else count

            elsecount = 0

        else
            'add volume
            stockvolume = stockvolume + ws.cells(i,7).value

            elsecount = elsecount + 1
        
            
        end if

        

    Next i

    


Next ws

end Sub



