<!DOCTYPE html>
<html lang="vi">
<head>
    <base href="<%=rsgp("website")%>/" />
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    
    <%
    ' =======================================================
    ' 1. KHỞI TẠO BIẾN SEO MẶC ĐỊNH TỪ TBLCONFIG
    ' =======================================================
    Dim seoTitle, seoDesc, seoImage, seoKeyword, siteName
    siteName   = rsgp("dongho")
    
    ' Gán giá trị mặc định cho trang chủ
    seoTitle   = siteName 
    seoDesc    = rsgp("mota")
    seoImage   = rsgp("anhdaidien")
    seoKeyword = rsgp("tukhoa")

    ' =======================================================
    ' 2. GHI ĐÈ GIÁ TRỊ SEO DỰA VÀO NGỮ CẢNH TRANG HIỆN TẠI
    ' =======================================================
    If gpdh = "1" Then
        ' Lưu ý: Nhomtin sẽ được bắt ở điều kiện bên dưới. Nếu gpdh=1 là trang Tin tức chung
        seoTitle = siteName & " | Tin tức"
    ElseIf gpdh = "2" Then
        seoTitle = siteName & " | Thư viện ảnh"
    ElseIf gpdh = "3" Then
        seoTitle = siteName & " | Thư viện video"
    ElseIf gpdh = "5" Then
        seoTitle = siteName & " | Phả đồ dọc"
    ElseIf gpdh = "6" Then
        seoTitle = siteName & " | Phả đồ theo đời"
    ElseIf gpdh = "7" Then
        seoTitle = siteName & " | Phả đồ ngang"
        
    ElseIf thongtin <> "" And IsNumeric(thongtin) Then
        Dim rsSEOThongTin
        Set rsSEOThongTin = conn.Execute("SELECT tieude FROM tblthongtin WHERE idthongtin = " & CLng(thongtin))
        If Not rsSEOThongTin.EOF Then seoTitle = siteName & " | " & rsSEOThongTin("tieude")
        rsSEOThongTin.Close : Set rsSEOThongTin = Nothing

    ElseIf nhomtin <> "" And IsNumeric(nhomtin) Then
        Dim rsSEONhomTin
        Set rsSEONhomTin = conn.Execute("SELECT tieude FROM tblnhomtin WHERE idnhomtin = " & CLng(nhomtin))
        If Not rsSEONhomTin.EOF Then seoTitle = siteName & " | " & rsSEONhomTin("tieude")
        rsSEONhomTin.Close : Set rsSEONhomTin = Nothing

    ElseIf giapha <> "" And IsNumeric(giapha) Then
        Dim rsSEOGiaPha
        Set rsSEOGiaPha = conn.Execute("SELECT hoten, gioitinh, hinhanh FROM tblgiapha WHERE idgiapha = " & CLng(giapha))
        If Not rsSEOGiaPha.EOF Then 
            Dim danhXung
            If CInt("0" & rsSEOGiaPha("gioitinh")) = 1 Then danhXung = "ông" Else danhXung = "bà" End If
            seoTitle = siteName & " | Thông tin " & danhXung & " " & rsSEOGiaPha("hoten")
            
            ' Nếu nhân vật có ảnh thì lấy ảnh nhân vật làm og:image, nếu không sẽ lấy mặc định anhdaidien
            If Not IsNull(rsSEOGiaPha("hinhanh")) And Trim(rsSEOGiaPha("hinhanh")) <> "" Then seoImage = rsSEOGiaPha("hinhanh")
        End If
        rsSEOGiaPha.Close : Set rsSEOGiaPha = Nothing

    ElseIf hinhanh <> "" And IsNumeric(hinhanh) Then
        Dim rsSEONhomAnh
        Set rsSEONhomAnh = conn.Execute("SELECT tieude, hinhanh FROM tblnhomanh WHERE idnhomanh = " & CLng(hinhanh))
        If Not rsSEONhomAnh.EOF Then 
            seoTitle = siteName & " | " & rsSEONhomAnh("tieude")
            If Not IsNull(rsSEONhomAnh("hinhanh")) And Trim(rsSEONhomAnh("hinhanh")) <> "" Then seoImage = rsSEONhomAnh("hinhanh")
        End If
        rsSEONhomAnh.Close : Set rsSEONhomAnh = Nothing

    ElseIf tintuc <> "" And IsNumeric(tintuc) Then
        Dim rsSEOTinTuc
        Set rsSEOTinTuc = conn.Execute("SELECT tieude, trichdan, hinhanh FROM tbltintuc WHERE idtintuc = " & CLng(tintuc))
        If Not rsSEOTinTuc.EOF Then 
            seoTitle = siteName & " | " & rsSEOTinTuc("tieude")
            If Not IsNull(rsSEOTinTuc("trichdan")) And Trim(rsSEOTinTuc("trichdan")) <> "" Then seoDesc = rsSEOTinTuc("trichdan")
            If Not IsNull(rsSEOTinTuc("hinhanh")) And Trim(rsSEOTinTuc("hinhanh")) <> "" Then seoImage = rsSEOTinTuc("hinhanh")
        End If
        rsSEOTinTuc.Close : Set rsSEOTinTuc = Nothing
    End If

    ' =======================================================
    ' 3. CHUẨN HÓA ĐƯỜNG DẪN ẢNH OPEN GRAPH TỐI ƯU CHO ZALO/FB
    ' =======================================================
    If seoImage <> "" And InStr(LCase(seoImage), "http") = 0 Then
        ' Loại bỏ ký tự ../ nếu có
        If Left(seoImage, 3) = "../" Then seoImage = Mid(seoImage, 4)
        ' Loại bỏ dấu gạch chéo đầu tiên để nối chuỗi không bị lỗi
        If Left(seoImage, 1) = "/" Then seoImage = Mid(seoImage, 2)
        
        Dim rootDomain : rootDomain = Trim(rsgp("website") & "")
        If Right(rootDomain, 1) <> "/" Then rootDomain = rootDomain & "/"
        
        seoImage = rootDomain & seoImage
    End If
    %>

    <!-- HIỂN THỊ CÁC THẺ SEO CHUẨN HÓA -->
    <title><%=seoTitle%></title>
    <meta name="description" content="<%=seoDesc%>">
    <meta name="keywords" content="<%=seoKeyword%>">
    <link rel="icon" href="<%=rsgp("icon")%>">

    <!-- Open Graph (OG) Meta Tags (Cho Facebook, Zalo, Viber...) -->
    <meta property="og:title" content="<%=seoTitle%>">
    <meta property="og:description" content="<%=seoDesc%>">
    <meta property="og:image" content="<%=seoImage%>">
    <meta property="og:type" content="website">
    <meta property="og:site_name" content="<%=siteName%>">
    <meta property="og:url" content="<%=rsgp("website") & Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")%>">

    <!-- FONT CHỮ SANG TRỌNG -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Italianno&family=Philosopher:ital,wght@0,400;0,700;1,400;1,700&family=Nunito:wght@400;600;700&display=swap" rel="stylesheet">
    
    <!-- THƯ VIỆN CSS/JS -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/slick-carousel/1.8.1/slick.min.css">
    
    <!-- KẾT NỐI CSS TĨNH MỚI TẠO (Đã thêm dấu / để chống lỗi đường dẫn khi dùng URL Rewrite) -->
    <link rel="stylesheet" href="/css/style.css?v=2.0">
    <% If bgChinhUrl <> "" Then %>
<style>
    :root {
        --bg-chinh-url: url('<%=bgChinhUrl%>');
    }
</style>
<% End If %>
    <!-- KHỞI TẠO BIẾN CSS & GIAO DIỆN HEADER GRADIENT ĐỒNG BỘ FOOTER -->
    <style>
        :root {
            --color-gold: #D1A659;
            --color-main: <%=rsgp("mausac")%>;
            --color-text: #333333;
            --color-bg-light: #F9F8F6;
            --font-heading: <%=rsgp("font_family")%>;
            --font-accent: 'Italianno', cursive;
            --font-body: <%=rsgp("font_family")%>;
            --font-size-base: <%=rsgp("font_size")%>;
            <% If Not IsNull(rsgp("bgchinh")) And Trim(rsgp("bgchinh")) <> "" Then %>
            --bg-chinh-url: url('<% If Left(Trim(rsgp("bgchinh")), 1) <> "/" And InStr(rsgp("bgchinh"), "http") = 0 Then Response.Write "/" End If %><%=Trim(rsgp("bgchinh"))%>');
            <% End If %>
        }

        body {
            font-size: var(--font-size-base);
            font-family: var(--font-body);
            color: var(--color-text);
        }

        /* --- 1. THANH HEADER: GRADIENT ĐA TẦNG VÀ HIỆU ỨNG CUỘN TRANG --- */
        #mainHeader {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            z-index: 99999;
            transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
            /* Gradient mờ tự nhiên từ trên xuống giúp chữ trắng luôn đọc tốt trên mọi ảnh nền slider/banner */
            background: linear-gradient(180deg, rgba(15, 23, 42, 0.75) 0%, rgba(15, 23, 42, 0) 100%);
        }

        /* Khi cuộn chuột: Biến hóa sang Gradient nhận diện thương hiệu đồng bộ Footer */
        #mainHeader.header-scrolled {
            background: linear-gradient(
                135deg, 
                color-mix(in srgb, var(--color-main) 88%, #0f172a 12%) 0%, 
                rgba(15, 23, 42, 0.96) 100%
            ) !important;
            box-shadow: 0 4px 25px rgba(0, 0, 0, 0.35);
            border-bottom: 2px solid var(--color-gold, #D1A659);
            backdrop-filter: blur(10px);
            -webkit-backdrop-filter: blur(10px);
        }

        /* --- 2. MENU ĐIỀU HƯỚNG CHÍNH: CHỮ TRẮNG SẮC NÉT --- */
        .main-menu {
            list-style: none;
            padding: 0;
            margin: 0;
            display: flex;
            align-items: center;
        }

        .main-menu > li {
            position: relative;
        }

        .main-menu > li > a {
            color: #ffffff !important; /* Trắng tinh khiết 100% */
            font-size: 17px;
            font-weight: 600;
            text-decoration: none;
            padding: 22px 16px;
            display: inline-flex;
            align-items: center;
            letter-spacing: 0.3px;
            transition: all 0.25s ease;
            text-shadow: 0 1px 3px rgba(0, 0, 0, 0.7); /* Chống lóa trên nền sáng */
        }

        .main-menu > li > a:hover,
        .main-menu > li:hover > a {
            color: var(--color-gold, #D1A659) !important;
            text-shadow: 0 0 8px rgba(209, 166, 89, 0.4);
        }

        .main-menu li a i.fa-angle-down {
            margin-left: 5px;
            font-size: 0.75em;
            color: var(--color-gold, #D1A659);
            transition: transform 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        }

        .main-menu li:hover > a i.fa-angle-down {
            transform: rotate(180deg);
        }

        /* Nút tìm kiếm trên thanh menu */
        .btn-search-icon {
            background: transparent;
            border: none;
            color: #ffffff;
            font-size: 16px;
            cursor: pointer;
            padding: 8px 12px;
            margin-left: 10px;
            transition: color 0.2s ease, transform 0.2s ease;
            text-shadow: 0 1px 3px rgba(0, 0, 0, 0.6);
        }
        .btn-search-icon:hover {
            color: var(--color-gold, #D1A659);
            transform: scale(1.1);
        }

        /* --- 3. MENU XỔ XUỐNG (DROPDOWN MENU): SANG TRỌNG & RÕ NÉT --- */
        .main-menu .dropdown-menu {
            position: absolute;
            top: 100%;
            left: 0;
            min-width: 230px;
            background: #ffffff;
            box-shadow: 0 12px 30px rgba(0, 0, 0, 0.2);
            border-top: 3px solid var(--color-gold, #D1A659);
            border-radius: 0 0 6px 6px;
            padding: 8px 0;
            list-style: none;
            opacity: 0;
            visibility: hidden;
            transform: translateY(12px);
            transition: all 0.3s cubic-bezier(0.165, 0.84, 0.44, 1);
            z-index: 9999;
        }

        .main-menu li:hover > .dropdown-menu {
            opacity: 1;
            visibility: visible;
            transform: translateY(0);
        }

        .main-menu .dropdown-menu li a {
            display: flex;
            align-items: center;
            gap: 8px;
            padding: 10px 20px;
            color: #1e293b !important;
            font-size: 16px;
            font-weight: 500;
            text-shadow: none;
            white-space: nowrap;
            transition: all 0.2s ease;
        }

        .main-menu .dropdown-menu li a i {
            color: #94a3b8;
            font-size: 11px;
            transition: all 0.2s ease;
        }

        .main-menu .dropdown-menu li a:hover {
            background-color: var(--color-bg-light, #f8fafc);
            color: var(--color-gold, #D1A659) !important;
            padding-left: 24px;
        }

        .main-menu .dropdown-menu li a:hover i {
            color: var(--color-gold, #D1A659);
        }
    </style>
</head>
<body>

    <!-- MODAL TÌM KIẾM TOÀN MÀN HÌNH -->
    <div class="search-modal" id="searchModal">
        <i class="fa fa-times search-close" onclick="toggleSearch()"></i>
        <div class="search-form">
            <h3>Tra cứu thông tin gia phả</h3>
            <!-- Đã chuyển hướng tìm kiếm về Phả đồ theo đời -->
<form action="/pha-do-theo-doi.html" method="get">
    <div class="search-input-wrap">
        <!-- Đổi name="keyword" thành name="tukhoa" cho khớp với biến m_tukhoa ở phả đồ -->
        <input type="text" name="tukhoa" placeholder="Nhập tên thành viên cần tìm..." autofocus>
        <button type="submit"><i class="fa fa-search"></i></button>
    </div>
</form>
        </div>
    </div>

    <!-- HEADER -->
    <header id="mainHeader">
        
        <!-- LỚP KÍNH MỜ NẰM CÙNG CẤP VỚI MENU ĐỂ ĐÈ NHAU -->
        <div class="mobile-overlay" id="mobileOverlay" onclick="toggleMobileMenu()"></div>

        <div class="container header-inner">
            <a href="<%=rsgp("website")%>" class="logo"><img src="<%=rsgp("logo")%>" alt="Logo"></a>
            
            <!-- NÚT MENU HAMBURGER BÊN PHẢI -->
            <button class="mobile-toggle" onclick="toggleMobileMenu()">
                <i class="fa fa-bars"></i>
            </button>

            <!-- KHỐI CHỨA MENU -->
            <div class="menu-wrapper" id="mobileMenu">
                <ul class="main-menu">
                    <li><a href="<%=rsgp("website")%>">Trang chủ</a></li>
                    
                    <!-- MENU GIỚI THIỆU ĐỘNG -->
                    <li>
                        <a href="#" onclick="toggleDropdown(event, this)">Giới thiệu <i class="fa fa-angle-down"></i></a>
                        <ul class="dropdown-menu">
                            <%
                                Dim rsMenuInfo
                                Set rsMenuInfo = Server.CreateObject("ADODB.Recordset")
                                rsMenuInfo.Open "SELECT idthongtin, tieude FROM tblthongtin ORDER BY idthongtin ASC", conn, 1, 1
                                Do While Not rsMenuInfo.EOF
                            %>
                            <!-- Đã áp dụng URL Rewrite SEO cho trang Thông tin -->
                            <li><a href="/thongtin/<%=rsMenuInfo("idthongtin")%>/<%=ConvertiTitoloInNomeScript(rsMenuInfo("tieude"))%>.html"><i class="fa fa-caret-right"></i> <%=rsMenuInfo("tieude")%></a></li>
                            <% 
                                rsMenuInfo.MoveNext
                                Loop
                                rsMenuInfo.Close : Set rsMenuInfo = Nothing
                            %>
                        </ul>
                    </li>
                    
                    <li>
                        <a href="/pha-do-doc.html" onclick="toggleDropdown(event, this)">Phả đồ <i class="fa fa-angle-down"></i></a>
                        <ul class="dropdown-menu">
                            <li><a href="/pha-do-doc.html"><i class="fa fa-caret-right"></i> Phả đồ dọc</a></li>
                            <li><a href="/pha-do-ngang.html"><i class="fa fa-caret-right"></i> Phả đồ ngang</a></li>
                            <li><a href="/pha-do-theo-doi.html"><i class="fa fa-caret-right"></i> Phả đồ theo đời</a></li>
                        </ul>
                    </li>                
                    
                    <!-- MENU TIN TỨC ĐỘNG -->
                    <%
                        Dim rsMenuNhomTin
                        Set rsMenuNhomTin = Server.CreateObject("ADODB.Recordset")
                        ' Lấy danh sách nhóm tin đã tạo trong trang quản trị
                        rsMenuNhomTin.Open "SELECT idnhomtin, tieude FROM tblnhomtin ORDER BY thutu ASC, idnhomtin DESC", conn, 1, 1
                        
                        If rsMenuNhomTin.EOF Then
                            ' TRƯỜNG HỢP 1: Dòng họ chưa tạo nhóm tin nào -> Chỉ hiển thị link thẳng
                    %>
                        <li><a href="/tin-tuc.html">Tin tức</a></li>
                    <%
                        Else
                            ' TRƯỜNG HỢP 2: Dòng họ có tạo nhóm tin -> Hiển thị menu xổ xuống
                    %>
                        <li>
                            <a href="/tin-tuc.html" onclick="toggleDropdown(event, this)">Tin tức <i class="fa fa-angle-down"></i></a>
                            <ul class="dropdown-menu">
                                <!-- Nút mặc định để xem toàn bộ tin tức -->                            
                                <%
                                    Do While Not rsMenuNhomTin.EOF
                                %>
                                <!-- Đã áp dụng URL Rewrite SEO cho Nhóm tin -->
                                <li><a href="/nhomtin/<%=rsMenuNhomTin("idnhomtin")%>/<%=ConvertiTitoloInNomeScript(rsMenuNhomTin("tieude"))%>.html"><i class="fa fa-caret-right"></i> <%=rsMenuNhomTin("tieude")%></a></li>
                                <% 
                                        rsMenuNhomTin.MoveNext
                                    Loop
                                %>
                            </ul>
                        </li>
                    <%
                        End If
                        
                        ' Đóng và giải phóng bộ nhớ
                        rsMenuNhomTin.Close 
                        Set rsMenuNhomTin = Nothing
                    %>                
                    <!-- MENU THƯ VIỆN GỘP -->
                    <li>
                        <a href="/thu-vien-anh.html" onclick="toggleDropdown(event, this)">Thư viện <i class="fa fa-angle-down"></i></a>
                        <ul class="dropdown-menu">
                            <!-- Đã sửa lại định dạng list danh mục chung -->
                            <li><a href="/thu-vien-anh.html"><i class="fa fa-caret-right"></i> Thư viện Hình ảnh</a></li>
                            <li><a href="/thu-vien-video.html"><i class="fa fa-caret-right"></i> Thư viện Video</a></li>
                        </ul>
                    </li>

                    <!-- MENU TÌM KIẾM TRÊN MOBILE -->
                        <li class="mobile-search-btn"><a href="javascript:void(0)" onclick="toggleSearch(); toggleMobileMenu();"><i class="fa fa-search"></i></a></li>
                </ul>
                
                <!-- ICON TÌM KIẾM TRÊN DESKTOP -->
                <button class="btn-search-icon" onclick="toggleSearch()" title="Tìm kiếm thành viên"><i class="fa fa-search"></i></button>
            </div>
        </div>
    </header>

    <!-- JAVASCRIPT XỬ LÝ SỰ KIỆN CUỘN TRANG (STICKY HEADER) -->
    <script>
        window.addEventListener('scroll', function() {
            const header = document.getElementById('mainHeader');
            if (window.scrollY > 50) {
                header.classList.add('header-scrolled');
            } else {
                header.classList.remove('header-scrolled');
            }
        });
    </script>
</body>
</html>