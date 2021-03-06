Attribute VB_Name = "Module1"
Option Explicit

Sub kmeans()
    Debug.Print Time
    
    Const Nclusters As Long = 20
    
    Const Nfigs As Long = 10000
    Const Ddim As Long = 28 ^ 2
    
    Dim data As Variant
    ReDim data(1 To Nfigs, 1 To Ddim) As Long
    Dim cluster(1 To Nfigs) As Long
    
    Dim fig As Long
    Dim c As Long, d As Long


    Randomize
    Application.ScreenUpdating = False
    
    
    '
    ' (1)
    '
    data = Sheets("mnist_data").Range("B1").Resize(Nfigs, Ddim)
    
    
    '
    ' (2)
    '
    For fig = 1 To Nfigs
        cluster(fig) = Int(Nclusters * Rnd) + 1
    Next fig

    
    Dim iteration As Long
    For iteration = 1 To 10
        Dim G As Variant, count As Variant
        ReDim G(1 To Nclusters, 1 To Ddim) As Long, count(1 To Nclusters) As Long 'G ÇÕçÇë¨âªÇÃÇΩÇﬂêÆêîå^Ç∆ÇµÇƒêÈåæÅi20Åì íˆìxèàóùéûä‘å∏Åj
        
        For fig = 1 To Nfigs
            c = cluster(fig)
            count(c) = count(c) + 1
        
            For d = 1 To Ddim
                G(c, d) = G(c, d) + data(fig, d)
            Next d
        Next fig
        
        For d = 1 To Ddim
            For c = 1 To Nclusters
                G(c, d) = G(c, d) \ count(c)
            Next c
        Next d
            
            

        Dim dist(1 To Nclusters) As LongLong
        Dim c_min As Long
        
        For fig = 1 To Nfigs

            For c = 1 To Nclusters
                dist(c) = 0
                For d = 1 To Ddim
                    dist(c) = dist(c) + (data(fig, d) - G(c, d)) ^ 2
                Next d
            Next c
                
            c_min = 1
            For c = 2 To Nclusters
                If dist(c) < dist(c_min) Then
                    c_min = c
                End If
            Next c
            
            cluster(fig) = c_min
        Next fig

    
    Next iteration
    

    '
    ' (3)
    '
    Dim i As Long, j As Long
    
    For c = 1 To Nclusters
        For i = 1 To 28
            For j = 1 To 28
                Sheets(1).Cells(28 * (c - 1) + i, j) = G(c, 28 * (i - 1) + j)
            Next j
        Next i
    Next c
    
    
    Debug.Print Time
End Sub






Sub load_test()
    Dim mnist
    Dim fig As Variant
    ReDim fig(1 To 28, 1 To 28)
    Dim i As Long, j As Long
    
    mnist = Sheets("mnist_data").Range("A1").CurrentRegion
    
    For i = 1 To 28
        For j = 1 To 28
            fig(i, j) = mnist(3, 1 + 28 * (i - 1) + j)
        Next j
    Next i
    
    
    Sheets(1).Range("A1").Resize(28, 28) = fig
    Stop
End Sub


