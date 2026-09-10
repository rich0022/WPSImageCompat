(function () {
  var paneKey = 'wps-image-compat-et.taskpane-id';
  function showToolbox() {
    var app = window.Application || window.wps;
    if (!app) { alert('WPS 加载项接口不可用。'); return false; }
    var id = app.PluginStorage && app.PluginStorage.getItem(paneKey);
    var pane = id && app.GetTaskPane ? app.GetTaskPane(id) : null;
    if (!pane) {
      pane = app.CreateTaskPane(GetUrlPath() + '/ui/taskpane.html', 'WPS Image Compat');
      if (!pane) { alert('无法打开工具箱。请检查加载项地址是否为 HTTPS 且可访问。'); return false; }
      if (app.PluginStorage) app.PluginStorage.setItem(paneKey, pane.ID);
    }
    pane.Visible = true;
    return true;
  }
  window.OnAddinLoad = function (ribbonUI) {
    var app = window.Application || window.wps;
    if (app && typeof app.ribbonUI !== 'object') app.ribbonUI = ribbonUI;
    return true;
  };
  window.OnAddInLoad = window.OnAddinLoad;
  window.OnAction = function (control) {
    if (control && control.Id === 'WPSImageCompat.Open') return showToolbox();
    return false;
  };
}());
