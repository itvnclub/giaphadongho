<%
' 1. XỬ LÝ AN TOÀN ĐƯỜNG DẪN ẢNH NỀN TỪ BGTREN
Dim bgFooterImg : bgFooterImg = ""
If Not IsNull(rsgp("bgtren")) And Trim(rsgp("bgtren")) <> "" Then
    bgFooterImg = Trim(rsgp("bgtren"))
    If Left(bgFooterImg, 1) <> "/" And InStr(bgFooterImg, "http") = 0 Then
        bgFooterImg = "/" & bgFooterImg
    End If
End If
%>

<!-- FOOTER ĐỒNG BỘ HOA VĂN BGTREN & OVERLAY -->
<footer class="footer" style="background-image: url('<%=bgFooterImg%>');">
    <!-- Lớp phủ màu nhận diện thương hiệu -->
    <div class="footer-overlay"></div>

    <div class="container footer-grid">
        
        <div class="foot-col" align="center">
            <img src="<%=rsgp("logof")%>" alt="Logo Footer" class="foot-logo">
        </div>

        <div class="foot-col">
            <h4 class="foot-title"><%=rsgp("dongho")%></h4>
            <ul class="foot-contact">
                <li><i class="fa fa-map-marker"></i> <span><%=rsgp("diachi")%></span></li>
                <li><i class="fa fa-user"></i> <span><b>Trưởng tộc:</b> <%=rsgp("toctruong")%> - <%=rsgp("dienthoai")%></span></li>
            </ul>
        </div>

        <div class="foot-col">
            <h4 class="foot-title">Tra cứu nhanh</h4>
            <div style="display: flex; gap: 30px;">
                <ul class="foot-links">
                    <%
                        Set rsMenuInfo = Server.CreateObject("ADODB.Recordset")
                        rsMenuInfo.Open "SELECT idthongtin, tieude FROM tblthongtin ORDER BY idthongtin ASC", conn, 1, 1
                        Do While Not rsMenuInfo.EOF
                    %>
                    <li><a href="/thongtin/<%=rsMenuInfo("idthongtin")%>/<%=ConvertiTitoloInNomeScript(rsMenuInfo("tieude"))%>.html"><%=rsMenuInfo("tieude")%></a></li>
                    <% 
                        rsMenuInfo.MoveNext
                        Loop
                        rsMenuInfo.Close : Set rsMenuInfo = Nothing
                    %>
                    <li><a href="/tin-tuc.html">Tin tức dòng họ</a></li>
                </ul>
                <ul class="foot-links">
                    <li><a href="/pha-do-doc.html">Phả đồ dọc</a></li>
                    <li><a href="/pha-do-ngang.html">Phả đồ ngang</a></li>
                    <li><a href="/pha-do-theo-doi.html">Phả đồ theo đời</a></li>
                    <li><a href="/thu-vien-anh.html">Thư viện hình ảnh</a></li>
                    <li><a href="/thu-vien-video.html">Thư viện video</a></li>
                </ul>
            </div>
        </div>

    </div>
    
    <div class="container copyright">
        &copy; <%=Year(Now())%> Hệ thống quản lý gia phả <b><%=rsgp("dongho")%></b>. Thiết kế và phát triển bởi GIAPHADONGHO.COM.VN
    </div>
</footer>

<!-- SCRIPT SCROLL & LUNAR JS -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.6.0/jquery.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/slick-carousel/1.8.1/slick.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/lunar-javascript/1.6.12/lunar.min.js"></script>
<script>
    function toggleMobileMenu() {
        var menu = document.getElementById('mobileMenu');
        var overlay = document.getElementById('mobileOverlay');
        var toggleBtn = document.querySelector('.mobile-toggle');
        var icon = toggleBtn.querySelector('i');
        
        menu.classList.toggle('active');
        overlay.classList.toggle('active');
        toggleBtn.classList.toggle('active');
        
        if(menu.classList.contains('active')) {
            icon.classList.remove('fa-bars');
            icon.classList.add('fa-times');
        } else {
            icon.classList.remove('fa-times');
            icon.classList.add('fa-bars');
        }
    }

    function toggleDropdown(e, element) {
        if(window.innerWidth <= 992) {
            e.preventDefault();
            var submenu = element.nextElementSibling;
            document.querySelectorAll('.dropdown-menu').forEach(function(menu) {
                if (menu !== submenu) menu.classList.remove('open');
            });
            submenu.classList.toggle('open');
        }
    }

    function toggleSearch() {
        var modal = document.getElementById('searchModal');
        modal.classList.toggle('active');
        if(modal.classList.contains('active')){
            setTimeout(() => { modal.querySelector('input').focus(); }, 300);
        }
    }

    $(document).ready(function(){
        $('#heroSlider').slick({
            slidesToShow: 1, 
            slidesToScroll: 1, 
            autoplay: true, 
            autoplaySpeed: 5000,
            fade: true,
            cssEase: 'linear',
            arrows: false, 
            dots: false,
            pauseOnHover: false
        });

        $('.news-slider').slick({
            slidesToShow: 3, slidesToScroll: 1, autoplay: true, autoplaySpeed: 3000,
            arrows: false, dots: true,
            responsive: [ { breakpoint: 992, settings: { slidesToShow: 2 } }, { breakpoint: 768, settings: { slidesToShow: 1 } } ]
        });
        
        if(window.innerWidth > 992) {
            $('.mobile-search-btn').hide();
        }
    });

    function openTab(evt, tabName) {
        $('.tab-content').hide();
        $('.tab-btn').removeClass('active');
        $('#' + tabName).fadeIn();
        $(evt.currentTarget).addClass('active');
    }

    var danhSachMat = [
    <%
        Set rsGioT = conn.Execute("SELECT idgiapha, hoten, nammat FROM tblgiapha WHERE nammat IS NOT NULL AND nammat <> '' AND loai_quanhe IN (1, 2) AND tinhtrang = 2")
        Do While Not rsGioT.EOF
            Response.Write "{ name: '" & Replace(rsGioT("hoten") & "", "'", "\'") & "', dateText: '" & Replace(Replace(rsGioT("nammat") & "", "'", "\'"), vbCrLf, " ") & "', link: '?giapha=" & rsGioT("idgiapha") & "' },"
            rsGioT.MoveNext
        Loop
        rsGioT.Close : Set rsGioT = Nothing
    %>
    ];

    var danhSachSinh = [
    <%
        Set rsSinhT = conn.Execute("SELECT idgiapha, hoten, namsinh FROM tblgiapha WHERE namsinh IS NOT NULL AND namsinh <> '' AND tinhtrang = 1")
        Do While Not rsSinhT.EOF
            Response.Write "{ name: '" & Replace(rsSinhT("hoten") & "", "'", "\'") & "', dateText: '" & Replace(Replace(rsSinhT("namsinh") & "", "'", "\'"), vbCrLf, " ") & "', link: '?giapha=" & rsSinhT("idgiapha") & "' },"
            rsSinhT.MoveNext
        Loop
        rsSinhT.Close : Set rsSinhT = Nothing
    %>
    ];

    function extractDayMonth(str) {
        if (!str) return null;
        var match = str.match(/(\d{1,2})\s*[\/\-\.]\s*(\d{1,2})/);
        if (match) return { day: parseInt(match[1], 10), month: parseInt(match[2], 10) };
        return null; 
    }

    function renderEvents(containerId, events, monthLabel) {
        var html = "";
        events.sort((a,b) => a.daysLeft - b.daysLeft).slice(0, 4).forEach(ev => {
            var status = ev.daysLeft === 0 ? "Hôm nay!" : "Còn " + ev.daysLeft + " ngày";
            html += `
            <div class="event-item">
                <div class="ev-date">
                    <div class="ev-day">${ev.day}</div>
                    <div class="ev-month">${monthLabel} ${ev.month}</div>
                </div>
                <div class="ev-info">
                    <h4><a href="${ev.person.link}">${ev.person.name}</a></h4>
                    <p><i class="fa fa-clock-o"></i> ${status}</p>
                </div>
            </div>`;
        });
        document.getElementById(containerId).innerHTML = html || "<p style='text-align:center; color:#fff;'>Không có sự kiện sắp tới.</p>";
    }

    setTimeout(() => {
        var today = new Date(); today.setHours(0,0,0,0);
        var evSinh = [];
        danhSachSinh.forEach(p => {
            var d = extractDayMonth(p.dateText);
            if(d) {
                var b = new Date(today.getFullYear(), d.month-1, d.day);
                var diff = Math.ceil((b - today)/86400000);
                if(diff < 0) { b.setFullYear(today.getFullYear()+1); diff = Math.ceil((b - today)/86400000); }
                evSinh.push({ person: p, day: d.day, month: d.month, daysLeft: diff });
            }
        });
        renderEvents('sinh-nhat-container', evSinh, "Tháng");

        if (typeof Lunar !== 'undefined') {
            var evGio = [];
            var tYear = Lunar.fromDate(today).getYear();
            danhSachMat.forEach(p => {
                var d = extractDayMonth(p.dateText);
                if(d) {
                    try {
                        var solar = Lunar.fromYmd(tYear, d.month, d.day).getSolar();
                        var jsDate = new Date(solar.getYear(), solar.getMonth()-1, solar.getDay());
                        var diff = Math.ceil((jsDate - today)/86400000);
                        if(diff < 0) {
                            solar = Lunar.fromYmd(tYear+1, d.month, d.day).getSolar();
                            jsDate = new Date(solar.getYear(), solar.getMonth()-1, solar.getDay());
                            diff = Math.ceil((jsDate - today)/86400000);
                        }
                        evGio.push({ person: p, day: d.day, month: d.month, daysLeft: diff });
                    }catch(e){}
                }
            });
            renderEvents('ngay-gio-container', evGio, "Tháng Âm");
        }
    }, 500);
</script>