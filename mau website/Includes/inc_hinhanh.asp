<%
' Hàm chuẩn hóa đường dẫn ảnh hiển thị ngoài trang chủ
Function GetPublicImgUrl(ByVal rawPath)
    rawPath = Trim(rawPath & "")
    If rawPath = "" Then
        GetPublicImgUrl = "/images/noimages.jpg"
        Exit Function
    End If
    If Left(LCase(rawPath), 4) = "http" Then
        GetPublicImgUrl = rawPath
        Exit Function
    End If
    rawPath = Replace(rawPath, "\", "/")
    Dim arrP, fileName
    If InStr(rawPath, "/") > 0 Then
        arrP = Split(rawPath, "/")
        fileName = arrP(UBound(arrP))
    Else
        fileName = rawPath
    End If
    If LCase(fileName) = "noimages.jpg" Then
        GetPublicImgUrl = "/images/noimages.jpg"
    Else
        GetPublicImgUrl = "/images/hinhanh/" & fileName
    End If
End Function

' --- 1. CẤU HÌNH PHÂN TRANG NHÓM ẢNH ---
reqPage = Request.QueryString("page")
pageSize = 10 ' Số lượng album/trang

If reqPage <> "" And IsNumeric(reqPage) Then
    currentPage = CInt(reqPage)
Else
    currentPage = 1
End If

' --- 2. TRUY VẤN DỮ LIỆU TỪ BẢNG TBLNHOMANH (KÈM ẢNH DỰ PHÒNG TỪ TBLANH) ---
Set rsAlbums = Server.CreateObject("ADODB.Recordset")
sqlAlbums = "SELECT n.idnhomanh, n.tieude, n.hinhanh, " & _
            "(SELECT TOP 1 a.hinhanh FROM tblanh a WHERE a.idnhomanh = n.idnhomanh ORDER BY a.thutu ASC, a.idanh DESC) AS anh_dau " & _
            "FROM tblnhomanh n ORDER BY ISNULL(n.thutu, 0) ASC, n.idnhomanh DESC"
rsAlbums.Open sqlAlbums, conn, 1, 1

If Not rsAlbums.EOF Then
    rsAlbums.PageSize = pageSize
    totalPages = rsAlbums.PageCount
    
    If currentPage > totalPages Then currentPage = totalPages
    If currentPage < 1 Then currentPage = 1
    
    rsAlbums.AbsolutePage = currentPage
End If
%>

<!-- ================= CSS GIAO DIỆN DANH SÁCH NHÓM ẢNH ================= -->
<style>
    /* BỐ CỤC FULL-WIDTH CHO ALBUM */
    .gallery-section { padding: 60px 0; background: #f8f9fa; }
    
    /* KHUNG DANH SÁCH ALBUM (2 NHÓM / HÀNG) */
    .album-grid {
        display: grid;
        grid-template-columns: repeat(2, 1fr);
        gap: 40px;
        margin-bottom: 40px;
    }

    /* CARD NHÓM ẢNH */
    .album-card {
        background: #fff; border-radius: 8px; overflow: hidden;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05); transition: transform 0.3s ease, box-shadow 0.3s ease;
        display: flex; flex-direction: column; border: 1px solid #f1f5f9; text-decoration: none;
    }
    .album-card:hover { transform: translateY(-5px); box-shadow: 0 10px 25px rgba(0,0,0,0.1); }
    
    .album-img-wrap {
        position: relative; width: 100%; padding-top: 70%; 
        overflow: hidden; background: #e2e8f0;
    }
    .album-img-wrap img {
        position: absolute; top: 0; left: 0; width: 100%; height: 100%;
        object-fit: cover; transition: transform 0.5s ease;
    }
    .album-card:hover .album-img-wrap img { transform: scale(1.08); }
    
    /* Lớp phủ (Overlay) Folder Icon */
    .album-overlay {
        position: absolute; inset: 0; background: rgba(0,0,0,0.4);
        display: flex; align-items: center; justify-content: center;
        opacity: 0; transition: 0.3s ease;
    }
    .album-card:hover .album-overlay { opacity: 1; }
    .album-overlay i {
        color: #fff; font-size: 40px; transform: scale(0.5); transition: 0.3s ease;
    }
    .album-card:hover .album-overlay i { transform: scale(1); }

    .album-info { padding: 15px 20px; text-align: center; background: #fff; }
    .album-title {
        font-family: var(--font-heading); font-size: 18px; font-weight: bold;
        color: #1e293b; margin: 0; transition: color 0.2s;
        white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
    }
    .album-card:hover .album-title { color: var(--color-main); }

    /* PHÂN TRANG */
    .pagination { display: flex; justify-content: center; flex-wrap: wrap; gap: 8px; margin-top: 20px; }
    .page-item {
        display: flex; align-items: center; justify-content: center; min-width: 40px; height: 40px;
        padding: 0 12px; background: #fff; color: #475569; text-decoration: none; border-radius: 4px;
        border: 1px solid #e2e8f0; font-weight: bold; transition: 0.2s;
    }
    .page-item:hover { background: #f8fafc; color: var(--color-main); border-color: var(--color-main); }
    .page-item.active { background: var(--color-main); color: #fff; border-color: var(--color-main); pointer-events: none; }
    .page-item.disabled { opacity: 0.5; pointer-events: none; }

    /* ĐÁP ỨNG THIẾT BỊ DI ĐỘNG */
    @media (max-width: 768px) {
        .album-grid { 
            grid-template-columns: 1fr; 
            gap: 20px; 
        }
        .album-title { font-size: 16px; }
    }
</style>

<!-- KHỐI BANNER CHUYÊN MỤC -->
<div class="page-banner" style="background-image: url('<%=bgTrenUrl%>');">
    <div class="banner-overlay"></div>
    <div class="container relative z-10 text-center">
        <h1 class="banner-title">Thư viện ảnh</h1>
        <p class="banner-breadcrumb">
            <a href="/">Trang chủ</a> <i class="fa fa-angle-right"></i> Thư viện ảnh
        </p>
    </div>
</div>

<section class="gallery-section">
    <div class="container">
        <!-- ================= KHỐI DANH SÁCH NHÓM ẢNH (FULL WIDTH) ================= -->
        <div class="album-grid">
            <%
            If Not rsAlbums.EOF Then
                count = 0
                Do While Not rsAlbums.EOF And count < pageSize
                    rawCover = Trim(rsAlbums("hinhanh") & "")
                    If rawCover = "" Then rawCover = Trim(rsAlbums("anh_dau") & "")
                    imgCover = GetPublicImgUrl(rawCover)
                    
                    ' ĐƯỜNG DẪN URL REWRITE CHUẨN SEO
                    Dim linkAlbumDetail
                    linkAlbumDetail = "/hinhanh/" & rsAlbums("idnhomanh") & "/" & ConvertiTitoloInNomeScript(rsAlbums("tieude")) & ".html"
            %>
            <a href="<%=linkAlbumDetail%>" class="album-card" title="<%=rsAlbums("tieude")%>">
                <div class="album-img-wrap">
                    <img src="<%=imgCover%>" alt="<%=rsAlbums("tieude")%>">
                    <div class="album-overlay">
                        <i class="fa fa-folder-open-o"></i>
                    </div>
                </div>
                <div class="album-info">
                    <h3 class="album-title"><%=rsAlbums("tieude")%></h3>
                </div>
            </a>
            <%
                    count = count + 1
                    rsAlbums.MoveNext
                Loop
            Else
            %>
            <div style="grid-column: 1 / -1; text-align: center; padding: 40px; background: #fff; border-radius: 8px; border: 1px solid #e2e8f0;">
                <p class="text-muted"><i class="fa fa-picture-o fa-3x" style="color: #cbd5e1; margin-bottom: 15px;"></i><br>Dòng họ chưa cập nhật nhóm ảnh nào.</p>
            </div>
            <%
            End If
            %>
        </div>

        <!-- ================= KHỐI PHÂN TRANG ================= -->
        <% If totalPages > 1 Then %>
        <div class="pagination">
            <% If currentPage > 1 Then %>
                <a href="/thu-vien-anh/trang-<%=currentPage - 1%>.html" class="page-item" title="Trang trước"><i class="fa fa-angle-double-left"></i></a>
            <% Else %>
                <span class="page-item disabled"><i class="fa fa-angle-double-left"></i></span>
            <% End If %>

            <% 
            For p = 1 To totalPages
                If p = currentPage Then
            %>
                    <span class="page-item active"><%=p%></span>
            <%  Else %>
                    <a href="/thu-vien-anh/trang-<%=p%>.html" class="page-item"><%=p%></a>
            <%  
                End If
            Next 
            %>

            <% If currentPage < totalPages Then %>
                <a href="/thu-vien-anh/trang-<%=currentPage + 1%>.html" class="page-item" title="Trang sau"><i class="fa fa-angle-double-right"></i></a>
            <% Else %>
                <span class="page-item disabled"><i class="fa fa-angle-double-right"></i></span>
            <% End If %>
        </div>
        <% End If %>
    </div>
</section>

<%
' Đóng kết nối
If IsObject(rsAlbums) Then
    rsAlbums.Close
    Set rsAlbums = Nothing
End If
%>
<script>document.body.classList.add('bg-full-watermark');</script>