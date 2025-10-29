Attribute VB_Name = "modHtmlMapa"
Public Function GetMapaHtml() As String
    Dim s As String
    s = s & "<!DOCTYPE html>" & vbCrLf
    s = s & "<html>" & vbCrLf
    s = s & "<head>" & vbCrLf
    s = s & "  <meta charset=""utf-8"">" & vbCrLf
    s = s & "  <title>Mapa VBA</title>" & vbCrLf
    s = s & "  <link rel=""stylesheet"" href=""https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"">" & vbCrLf
    s = s & "  <script src=""https://unpkg.com/leaflet@1.9.4/dist/leaflet.js""></script>" & vbCrLf
    s = s & "  <style>html, body, #map { height: 100%; margin: 0; padding: 0; }</style>" & vbCrLf
    s = s & "</head>" & vbCrLf
    s = s & "<body>" & vbCrLf
    s = s & "  <div id=""map""></div>" & vbCrLf
    s = s & "  <script>" & vbCrLf
    s = s & "    var map = L.map('map').setView([-34.6037, -58.3816], 12);" & vbCrLf
    s = s & "    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {maxZoom: 19}).addTo(map);" & vbCrLf
    s = s & "    var marker;" & vbCrLf
    s = s & "    map.on('click', function (e) {" & vbCrLf
    s = s & "      var lat = e.latlng.lat; var lng = e.latlng.lng;" & vbCrLf
    s = s & "      if (marker) map.removeLayer(marker);" & vbCrLf
    s = s & "      marker = L.marker([lat, lng]).addTo(map);" & vbCrLf
    s = s & "      reportCoordsString(lat, lng);" & vbCrLf
    s = s & "    });" & vbCrLf
    s = s & "    function reportCoordsString(lat, lng) {" & vbCrLf
    s = s & "      var s = lat + ',' + lng; document.title = 'coords_str:' + s;" & vbCrLf
    s = s & "    }" & vbCrLf
    s = s & "    function setMarkerFromHost(lat, lng) {" & vbCrLf
    s = s & "      if (marker) map.removeLayer(marker);" & vbCrLf
    s = s & "      marker = L.marker([lat, lng]).addTo(map);" & vbCrLf
    s = s & "      try { map.setView([lat, lng], map.getZoom()); } catch (e) {}" & vbCrLf
    s = s & "    }" & vbCrLf
    s = s & "    function setMarkerFromHostText(t) {" & vbCrLf
    s = s & "      if (!t) return;" & vbCrLf
    s = s & "      var sep = t.indexOf(';') !== -1 ? ';' : ',';" & vbCrLf
    s = s & "      var parts = t.split(sep);" & vbCrLf
    s = s & "      if (parts.length !== 2) return;" & vbCrLf
    s = s & "      var lat = parseFloat(String(parts[0]).replace(',', '.'));" & vbCrLf
    s = s & "      var lng = parseFloat(String(parts[1]).replace(',', '.'));" & vbCrLf
    s = s & "      if (isFinite(lat) && isFinite(lng)) setMarkerFromHost(lat, lng);" & vbCrLf
    s = s & "    }" & vbCrLf
    s = s & "    (function(){" & vbCrLf
    s = s & "      try {" & vbCrLf
    s = s & "        var s = '';" & vbCrLf
    s = s & "        if (location.search) {" & vbCrLf
    s = s & "          var m = /[?&]coords=([^&]+)/i.exec(location.search);" & vbCrLf
    s = s & "          if (m && m[1]) s = decodeURIComponent(m[1]);" & vbCrLf
    s = s & "        }" & vbCrLf
    s = s & "        if (!s && location.hash) {" & vbCrLf
    s = s & "          var h = location.hash.substring(1);" & vbCrLf
    s = s & "          if (h.toLowerCase().indexOf('coords:') === 0) s = h.substring(7);" & vbCrLf
    s = s & "        }" & vbCrLf
    s = s & "        if (s) setMarkerFromHostText(s);" & vbCrLf
    s = s & "      } catch (e) {}" & vbCrLf
    s = s & "    })();" & vbCrLf
    s = s & "  </script>" & vbCrLf
    s = s & "</body>" & vbCrLf
    s = s & "</html>" & vbCrLf
    GetMapaHtml = s
End Function
