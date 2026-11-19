<%
idgiapha = Request.QueryString("giapha")

If idgiapha <> "" And IsNumeric(idgiapha) Then
    ' Lấy thông tin nhân vật chính và tên đời
    sqlDetail = "SELECT g.*, d.tieude AS Tendoi FROM tblgiapha g LEFT JOIN tbldoi d ON g.iddoi = d.iddoi WHERE g.idgiapha = " & idgiapha & " AND g.hienthi = 1"
    Set rsDetail = conn.Execute(sqlDetail)
    
    If Not rsDetail.EOF Then
        ' --- 1. HÀM HỖ TRỢ XỬ LÝ TRƯỜNG RỖNG ---
        Function CheckEmpty(val)
            If IsNull(val) Or Trim(val) = "" Then
                CheckEmpty = "<span style='color: #94a3b8; font-style: italic;'>Đang cập nhật</span>"
            Else
                CheckEmpty = val
            End If
        End Function

        ' --- 2. PHÂN TÍCH THỨ BẬC (Con thứ mấy, con của ai) ---
        Dim tenCha, strThuTu
        tenCha = "Đang cập nhật"
        
        If rsDetail("idcha") > 0 Then
            Dim rsCha
            Set rsCha = conn.Execute("SELECT hoten FROM tblgiapha WHERE idgiapha = " & rsDetail("idcha"))
            If Not rsCha.EOF Then tenCha = rsCha("hoten")
            rsCha.Close : Set rsCha = Nothing
        End If
        
        If rsDetail("idcha") = 0 Then
            strThuTu = "Thủy tổ dòng họ"
        Else
            If Not IsNull(rsDetail("thutu")) And rsDetail("thutu") > 0 Then
                strThuTu = "Con thứ " & rsDetail("thutu") & " của ông <strong>" & tenCha & "</strong>"
            Else
                strThuTu = "Con của ông <strong>" & tenCha & "</strong>"
            End If
        End If

        ' --- 3. XỬ LÝ HÌNH ẢNH MẶC ĐỊNH ---
        Dim hinhAnh : hinhAnh = rsDetail("hinhanh") & ""
        If hinhAnh = "" Then
            If rsDetail("gioitinh") = 1 Then
                hinhAnh = "Images/Hinhanh/default_male.jpg" 
            Else
                hinhAnh = "Images/Hinhanh/default_female.jpg" 
            End If
        End If
%>

<!-- ================= CSS GIAO DIỆN CHI TIẾT ĐỒNG BỘ VỚI INC_VIDEO ================= -->
<style>
    /* KHỐI NỘI DUNG CHI TIẾT */
    .gp-section { padding: 60px 0; background: #f8f9fa; }
    .gp-card { background: #fff; border-radius: 8px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); border: 1px solid #f1f5f9; padding: 40px; margin-bottom: 40px; }
    .gp-profile-wrapper { display: flex; flex-wrap: wrap; gap: 40px; }
    
    .gp-avatar { flex: 0 0 250px; text-align: center; }
    .gp-avatar img { width: 100%; border-radius: 8px; border: 3px solid #e2e8f0; object-fit: cover; }
    
    .gp-info { flex: 1; min-width: 300px; }
    .gp-name { font-family: var(--font-heading); font-size: 32px; color: var(--color-main); margin: 0 0 8px 0; }
    .gp-hierarchy { font-size: 16px; color: #475569; margin-bottom: 25px; padding-bottom: 15px; border-bottom: 1px dashed #cbd5e1; }
    
    .gp-info-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 15px 30px; }
    .gp-info-item { display: flex; align-items: baseline; }
    .gp-info-label { font-weight: bold; color: #334155; width: 130px; flex-shrink: 0; }
    .gp-info-val { color: #0f172a; flex: 1; }
    
    .gp-heading { font-family: var(--font-heading); color: var(--color-main); border-bottom: 2px solid var(--color-gold); padding-bottom: 10px; margin-top: 0; margin-bottom: 25px; font-size: 22px; display: inline-block; }
    
    /* LƯỚI NGƯỜI THÂN (VỢ/CHỒNG, CON CÁI) */
    .gp-family-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(320px, 1fr)); gap: 20px; }
    .gp-family-card { display: flex; align-items: center; gap: 15px; background: #f8fafc; padding: 15px; border-radius: 8px; border: 1px solid #e2e8f0; text-decoration: none; transition: 0.3s ease; }
    .gp-family-card:hover { transform: translateY(-3px); box-shadow: 0 10px 20px rgba(0,0,0,0.05); border-color: var(--color-main); background: #fff; }
    .gp-family-icon { width: 45px; height: 45px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 20px; color: #fff; flex-shrink: 0; }
    .gp-family-icon.male { background: var(--color-main, #046c52); }
    .gp-family-icon.female { background: #d81b60; }
    .gp-family-details h4 { margin: 0 0 5px 0; color: #1e293b; font-size: 16px; font-weight: bold; transition: 0.2s; }
    .gp-family-card:hover .gp-family-details h4 { color: var(--color-main); }
    .gp-family-details p { margin: 0; font-size: 13px; color: #64748b; }

    @media (max-width: 768px) {
        .gp-card { padding: 20px; }
        .gp-info-grid { grid-template-columns: 1fr; }
        .gp-avatar { flex: 1 1 100%; }
        .gp-info-item { flex-direction: column; gap: 4px; }
        .gp-info-label { width: 100%; }
    }
</style>

<!-- KHỐI BANNER -->
<div class="page-banner" style="background-image: url('<%=bgTrenUrl%>');">
    <div class="banner-overlay"></div>
    <div class="container relative z-10 text-center">
        <h1 class="banner-title"><%=rsDetail("hoten")%></h1>
        <p class="banner-breadcrumb">
            <a href="/">Trang chủ</a> <i class="fa fa-angle-right"></i> 
            <a href="/gia-pha.html">Gia phả</a> <i class="fa fa-angle-right"></i> 
            Chi tiết nhân vật
        </p>
    </div>
</div>

<section class="gp-section">
    <div class="container">
        
        <!-- ================= THÔNG TIN CÁ NHÂN ================= -->
        <div class="gp-card">
            <div class="gp-profile-wrapper">
                <div class="gp-avatar">
                    <img src="<%=hinhAnh%>" alt="<%=rsDetail("hoten")%>" />
                </div>
                
                <div class="gp-info">
                    <h2 class="gp-name"><%=rsDetail("hoten")%></h2>
                    <div class="gp-hierarchy">
                        <%=strThuTu%> 
                        <% If Not IsNull(rsDetail("Tendoi")) Then %>
                            (Đời thứ <strong><%=rsDetail("Tendoi")%></strong>)
                        <% End If %>
                    </div>
                    
                    <div class="gp-info-grid">
                        <div class="gp-info-item">
                            <span class="gp-info-label">Giới tính:</span>
                            <span class="gp-info-val"><% If rsDetail("gioitinh") = 1 Then Response.Write("Nam") Else Response.Write("Nữ") End If %></span>
                        </div>
                        <div class="gp-info-item">
                            <span class="gp-info-label">Tên tự/Hiệu:</span>
                            <span class="gp-info-val"><%=CheckEmpty(rsDetail("tentu"))%></span>
                        </div>
                        <div class="gp-info-item">
                            <span class="gp-info-label">Năm sinh:</span>
                            <span class="gp-info-val"><%=CheckEmpty(rsDetail("namsinh"))%></span>
                        </div>
                        <div class="gp-info-item">
                            <span class="gp-info-label">Nghề nghiệp:</span>
                            <span class="gp-info-val"><%=CheckEmpty(rsDetail("nghenghiep"))%></span>
                        </div>
                        <div class="gp-info-item" style="grid-column: 1 / -1;">
                            <span class="gp-info-label">Địa chỉ:</span>
                            <span class="gp-info-val"><%=CheckEmpty(rsDetail("diachi"))%></span>
                        </div>
                        
                        <!-- CHỈ HIỂN THỊ CÁC TRƯỜNG NÀY NẾU ĐÃ MẤT -->
                        <% If rsDetail("tinhtrang") = 2 Then %>
                        <div class="gp-info-item">
                            <span class="gp-info-label">Tình trạng:</span>
                            <span class="gp-info-val" style="color: #ef4444; font-weight: bold;">Đã mất</span>
                        </div>
                        <div class="gp-info-item">
                            <span class="gp-info-label">Năm mất:</span>
                            <span class="gp-info-val"><%=CheckEmpty(rsDetail("nammat"))%></span>
                        </div>
                        <div class="gp-info-item">
                            <span class="gp-info-label">Hưởng thọ:</span>
                            <span class="gp-info-val"><% If Not IsNull(rsDetail("huongtho")) And Trim(rsDetail("huongtho")) <> "" Then Response.Write(rsDetail("huongtho") & " tuổi") Else Response.Write(CheckEmpty("")) End If %></span>
                        </div>
                        <div class="gp-info-item" style="grid-column: 1 / -1;">
                            <span class="gp-info-label">Nơi an táng:</span>
                            <span class="gp-info-val"><%=CheckEmpty(rsDetail("noiantang"))%></span>
                        </div>
                        <% End If %>
                    </div>
                </div>
            </div>
        </div>

        <!-- ================= TIỂU SỬ ================= -->
        <% If Not IsNull(rsDetail("tieusu")) And Trim(rsDetail("tieusu")) <> "" Then %>
        <div class="gp-card">
            <h3 class="gp-heading">Tiểu sử & Sự nghiệp</h3>
            <div style="line-height: 1.7; color: #334155; text-align: justify;">
                <%=rsDetail("tieusu")%>
            </div>
        </div>
        <% End If %>
        
        <!-- ================= THÔNG TIN VỢ / CHỒNG ================= -->
        <%
            Dim sqlVoChong
            ' Truy vấn 2 chiều: 
            ' 1. Tìm người có id_vochong = idgiapha hiện tại
            ' 2. Nếu idgiapha hiện tại là vợ/chồng (có id_vochong > 0), thì tìm người gốc dựa trên id_vochong đó.
            sqlVoChong = "SELECT idgiapha, hoten, gioitinh, namsinh, tinhtrang FROM tblgiapha WHERE id_vochong = " & idgiapha & " OR (idgiapha = (SELECT TOP 1 id_vochong FROM tblgiapha WHERE idgiapha = " & idgiapha & " AND id_vochong > 0)) ORDER BY idgiapha ASC"
            
            Dim rsVC
            Set rsVC = conn.Execute(sqlVoChong)
            
            If Not rsVC.EOF Then
        %>
        <div class="gp-card">
            <h3 class="gp-heading">Thông tin Vợ / Chồng</h3>
            <div class="gp-family-grid">
                <% Do While Not rsVC.EOF 
                    Dim iconVC
                    If rsVC("gioitinh") = 1 Then iconVC = "male" Else iconVC = "female" End If
                %>
                <a href="/giapha/<%=rsVC("idgiapha")%>" class="gp-family-card">
                    <div class="gp-family-icon <%=iconVC%>"><i class="fa fa-user"></i></div>
                    <div class="gp-family-details">
                        <h4><%=rsVC("hoten")%></h4>
                        <p>
                            Năm sinh: <%=CheckEmpty(rsVC("namsinh"))%>
                            <% If rsVC("tinhtrang") = 2 Then Response.Write(" - <span style='color:#ef4444;'>(Đã mất)</span>") End If %>
                        </p>
                    </div>
                </a>
                <% rsVC.MoveNext : Loop %>
            </div>
        </div>
        <%
            End If
            rsVC.Close : Set rsVC = Nothing
        %>

        <!-- ================= DANH SÁCH CON CÁI ================= -->
        <%
            Dim rsCon
            ' Lấy danh sách con dựa vào idcha hoặc idme
            Set rsCon = conn.Execute("SELECT idgiapha, hoten, thutu, gioitinh, namsinh, tinhtrang FROM tblgiapha WHERE idcha = " & idgiapha & " OR idme = " & idgiapha & " ORDER BY thutu ASC, namsinh ASC")
            
            If Not rsCon.EOF Then
        %>
        <div class="gp-card">
            <h3 class="gp-heading">Danh sách Hậu duệ (Các con)</h3>
            <div class="gp-family-grid">
                <% Do While Not rsCon.EOF 
                    Dim iconCon, txtGioiTinh, txtThuTu
                    If rsCon("gioitinh") = 1 Then 
                        iconCon = "male"
                        txtGioiTinh = "Con trai"
                    Else 
                        iconCon = "female"
                        txtGioiTinh = "Con gái"
                    End If
                    
                    If Not IsNull(rsCon("thutu")) And rsCon("thutu") > 0 Then
                        txtThuTu = " (Con thứ " & rsCon("thutu") & ")"
                    Else
                        txtThuTu = ""
                    End If
                %>
                <a href="/giapha/<%=rsCon("idgiapha")%>" class="gp-family-card">
                    <div class="gp-family-icon <%=iconCon%>"><i class="fa fa-user"></i></div>
                    <div class="gp-family-details">
                        <h4><%=rsCon("hoten")%> <small style="color:var(--color-main); font-weight:normal;"><%=txtThuTu%></small></h4>
                        <p>
                            <%=txtGioiTinh%> - Năm sinh: <%=CheckEmpty(rsCon("namsinh"))%>
                            <% If rsCon("tinhtrang") = 2 Then Response.Write(" - <span style='color:#ef4444;'>(Đã mất)</span>") End If %>
                        </p>
                    </div>
                </a>
                <% rsCon.MoveNext : Loop %>
            </div>
        </div>
        <%
            End If
            rsCon.Close : Set rsCon = Nothing
        %>

    </div>
</section>

<%
    End If
    rsDetail.Close : Set rsDetail = Nothing
Else
    Response.Write("<div class='container' style='padding: 100px 0; text-align: center;'><h3 style='color: red;'>Dữ liệu không hợp lệ hoặc nhân vật không tồn tại!</h3></div>")
End If
%>
<script>document.body.classList.add('bg-full-watermark');</script>