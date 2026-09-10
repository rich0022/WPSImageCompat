function GetUrlPath() {
  var url = decodeURI(document.location.toString());
  return url.substring(0, url.lastIndexOf('/'));
}
