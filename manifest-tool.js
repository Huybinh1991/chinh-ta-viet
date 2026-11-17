/*
 * manifest-tool.js — Tạo sẵn file manifest.xml với đúng địa chỉ trang add-in của bạn.
 * Mở trang add-in (VD: https://ten-ban.github.io/chinh-ta-viet/) bằng trình duyệt, bấm nút
 * "Tải file manifest.xml" là có file đã điền sẵn, không cần sửa bằng Notepad.
 */
(function (root) {
  "use strict";
  var TEMPLATE = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<!-- manifest.xml cho add-in Chính Tả Việt — được tạo tự động từ trang add-in -->\n<OfficeApp xmlns=\"http://schemas.microsoft.com/office/appforoffice/1.1\"\n           xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"\n           xmlns:bt=\"http://schemas.microsoft.com/office/officeappbasictypes/1.0\"\n           xmlns:ov=\"http://schemas.microsoft.com/office/taskpaneappversionoverrides\"\n           xsi:type=\"TaskPaneApp\">\n  <Id>47245177-f1ed-452f-94a0-f8b3e1e161d8</Id>\n  <Version>1.0.0.0</Version>\n  <ProviderName>Chinh Ta Viet</ProviderName>\n  <DefaultLocale>vi-VN</DefaultLocale>\n  <DisplayName DefaultValue=\"Chính Tả Việt\"/>\n  <Description DefaultValue=\"Kiểm tra và sửa lỗi chính tả tiếng Việt bằng luật chính tả và AI Claude.\"/>\n  <IconUrl DefaultValue=\"https://TEN-GITHUB.github.io/chinh-ta-viet/icon-32.png\"/>\n  <HighResolutionIconUrl DefaultValue=\"https://TEN-GITHUB.github.io/chinh-ta-viet/icon-64.png\"/>\n  <SupportUrl DefaultValue=\"https://TEN-GITHUB.github.io/chinh-ta-viet/index.html\"/>\n  <Hosts>\n    <Host Name=\"Document\"/>\n  </Hosts>\n  <Requirements>\n    <Sets DefaultMinVersion=\"1.1\">\n      <Set Name=\"WordApi\" MinVersion=\"1.1\"/>\n    </Sets>\n  </Requirements>\n  <DefaultSettings>\n    <SourceLocation DefaultValue=\"https://TEN-GITHUB.github.io/chinh-ta-viet/index.html\"/>\n  </DefaultSettings>\n  <Permissions>ReadWriteDocument</Permissions>\n\n  <VersionOverrides xmlns=\"http://schemas.microsoft.com/office/taskpaneappversionoverrides\" xsi:type=\"VersionOverridesV1_0\">\n    <Hosts>\n      <Host xsi:type=\"Document\">\n        <DesktopFormFactor>\n          <GetStarted>\n            <Title resid=\"GetStarted.Title\"/>\n            <Description resid=\"GetStarted.Desc\"/>\n            <LearnMoreUrl resid=\"Taskpane.Url\"/>\n          </GetStarted>\n          <FunctionFile resid=\"Commands.Url\"/>\n          <ExtensionPoint xsi:type=\"PrimaryCommandSurface\">\n            <OfficeTab id=\"TabHome\">\n              <Group id=\"ChinhTaViet.Group\">\n                <Label resid=\"Group.Label\"/>\n                <Icon>\n                  <bt:Image size=\"16\" resid=\"Icon.16\"/>\n                  <bt:Image size=\"32\" resid=\"Icon.32\"/>\n                  <bt:Image size=\"80\" resid=\"Icon.80\"/>\n                </Icon>\n                <Control xsi:type=\"Button\" id=\"ChinhTaViet.Open\">\n                  <Label resid=\"Button.Label\"/>\n                  <Supertip>\n                    <Title resid=\"Button.Label\"/>\n                    <Description resid=\"Button.Tip\"/>\n                  </Supertip>\n                  <Icon>\n                    <bt:Image size=\"16\" resid=\"Icon.16\"/>\n                    <bt:Image size=\"32\" resid=\"Icon.32\"/>\n                    <bt:Image size=\"80\" resid=\"Icon.80\"/>\n                  </Icon>\n                  <Action xsi:type=\"ShowTaskpane\">\n                    <TaskpaneId>ChinhTaVietPane</TaskpaneId>\n                    <SourceLocation resid=\"Taskpane.Url\"/>\n                  </Action>\n                </Control>\n              </Group>\n            </OfficeTab>\n          </ExtensionPoint>\n        </DesktopFormFactor>\n      </Host>\n    </Hosts>\n    <Resources>\n      <bt:Images>\n        <bt:Image id=\"Icon.16\" DefaultValue=\"https://TEN-GITHUB.github.io/chinh-ta-viet/icon-16.png\"/>\n        <bt:Image id=\"Icon.32\" DefaultValue=\"https://TEN-GITHUB.github.io/chinh-ta-viet/icon-32.png\"/>\n        <bt:Image id=\"Icon.80\" DefaultValue=\"https://TEN-GITHUB.github.io/chinh-ta-viet/icon-80.png\"/>\n      </bt:Images>\n      <bt:Urls>\n        <bt:Url id=\"Commands.Url\" DefaultValue=\"https://TEN-GITHUB.github.io/chinh-ta-viet/commands.html\"/>\n        <bt:Url id=\"Taskpane.Url\" DefaultValue=\"https://TEN-GITHUB.github.io/chinh-ta-viet/index.html\"/>\n      </bt:Urls>\n      <bt:ShortStrings>\n        <bt:String id=\"GetStarted.Title\" DefaultValue=\"Chính Tả Việt đã sẵn sàng\"/>\n        <bt:String id=\"Group.Label\" DefaultValue=\"Chính Tả Việt\"/>\n        <bt:String id=\"Button.Label\" DefaultValue=\"Kiểm tra chính tả\"/>\n      </bt:ShortStrings>\n      <bt:LongStrings>\n        <bt:String id=\"GetStarted.Desc\" DefaultValue=\"Bấm nút Kiểm tra chính tả trên tab Home để mở khung kiểm tra lỗi.\"/>\n        <bt:String id=\"Button.Tip\" DefaultValue=\"Mở khung kiểm tra và sửa lỗi chính tả tiếng Việt.\"/>\n      </bt:LongStrings>\n    </Resources>\n  </VersionOverrides>\n</OfficeApp>\n";
  var PLACEHOLDER = "https://TEN-GITHUB.github.io/chinh-ta-viet/";

  function baseUrl(href) {
    var u = String(href).split("#")[0].split("?")[0];
    return u.slice(0, u.lastIndexOf("/") + 1);
  }
  function build(href) { return TEMPLATE.split(PLACEHOLDER).join(baseUrl(href)); }

  function download(href) {
    var xml = build(href);
    var blob = new Blob([xml], { type: "application/xml" });
    if (window.navigator && window.navigator.msSaveBlob) { window.navigator.msSaveBlob(blob, "manifest.xml"); return; }
    var a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "manifest.xml";
    document.body.appendChild(a);
    a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.parentNode.removeChild(a); }, 1000);
  }

  // Gắn nút vào trang (chỉ hiện khi trang đã ở trên mạng với https, mở ngoài Word)
  function init() {
    var box = document.getElementById("manifestBox");
    if (!box) return;
    var ok = location.protocol === "https:";
    box.style.display = "block";
    var btn = document.getElementById("btnManifest");
    var note = document.getElementById("manifestNote");
    if (!ok) {
      btn.disabled = true;
      note.textContent = "Nút này chỉ dùng được khi trang đã được đưa lên GitHub Pages (địa chỉ bắt đầu bằng https://).";
      return;
    }
    note.textContent = "Địa chỉ add-in sẽ ghi vào file: " + baseUrl(location.href);
    btn.onclick = function () { download(location.href); };
  }

  var api = { build: build, baseUrl: baseUrl, init: init };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.ManifestTool = api;
})(this);
