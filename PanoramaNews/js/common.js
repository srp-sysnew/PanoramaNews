(function(g){
"use strict";

// MyDate（Dateに関数追加)
g.createMyDate = function() {

    var _obj = null;
    
    if (arguments.length==1) {
        var _str = arguments[0];    
        if (typeof _str === 'string') {
            if (_str.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})Z$/i)) {
                var _obj = new Date(
                    Date.UTC(parseInt(RegExp.$1,10), parseInt(RegExp.$2,10)-1, parseInt(RegExp.$3,10), parseInt(RegExp.$4,10), parseInt(RegExp.$5,10), parseInt(RegExp.$6,10))
                    );
//            } else if (_str.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})Z$/i)) {
//                var _obj = new Date(
//                    Date.UTC(parseInt(RegExp.$1,10), parseInt(RegExp.$2,10)-1, parseInt(RegExp.$3,10), parseInt(RegExp.$4,10), parseInt(RegExp.$5,10), parseInt(RegExp.$6,10))
//                    );
                    
            } else {
                _obj = new Date(_str);
            }
        }
    } else if (arguments.length == 3) {
        // 年月日対応
        _obj = new Date(arguments[0],arguments[1],arguments[2]);
    }        
    if (_obj === null) {
        _obj = new Date();
    }

    _obj.youbi=['日','月', '日', '水', '木', '金', '土'];

    _obj.toDisplayTitle = function () {
        var _y=_obj.getFullYear(), _h = _obj.getHours(), _m = _obj.getMinutes();
        return _y+ '年' + '(平成' + (_y-1988) + '年)' + (_obj.getMonth()+1) + '月' + (_obj.getDate()) + '日 '+ _obj.youbi[_obj.getDay()] + '曜日' 
            + ' ' + ((_h < 10) ? '0' : '') + _h + ':' + ((_m < 10) ? '0' : '') + _m;
    };

    _obj.toDisplayShort = function(){
        return (_obj.getMonth() + 1) + '月' + (_obj.getDate()) + '日(' + _obj.youbi[_obj.getDay()] + ')';
    };

    _obj.toDisplay = function() {
        var _h = _obj.getHours(), _m = _obj.getMinutes();
        return _obj.toDisplayShort() + ' ' + ((_h < 10) ? '0' : '') + _h + ':' + ((_m < 10) ? '0' : '') + _m;
    };

    _obj.toYmd = function() {
        var yyyy = _obj.getFullYear(), mm = _obj.getMonth() + 1, dd = _obj.getDate();

        if (mm < 10) { mm = "0" + mm; }
        if (dd < 10) { dd = "0" + dd; }
        return yyyy + '-' + mm + '-' + dd;
    };
    
    _obj.addTime = function(n) {
        var _n = _obj.getTime() +  n;
        var _res = createMyDate();
        _res.setTime(_n);
        return _res;
    };

    _obj.toStringForApi = function() {
        var _y = _obj.getUTCFullYear(), _m = _obj.getUTCMonth()+1, _d = _obj.getUTCDate(), 
        _h = _obj.getUTCHours(), _i = _obj.getUTCMinutes(), _s = _obj.getUTCSeconds();
        
        if (_m < 10) { _m = '0' + _m; }
        if (_d < 10) { _d = '0' + _d; }
        if (_h < 10) { _h = '0' + _h; }
        if (_i < 10) { _i = '0' + _i; }
        if (_s < 10) { _s = '0' + _s; }
        
        return _y + '-' + _m + '-' + _d + 'T' + _h + ':' + _i + ':' + _s + 'Z';
    };

    return _obj;
};

g.createToday = function(){
    var _dt = new Date();
    return createMyDate(_dt.getFullYear(), _dt.getMonth(), _dt.getDate());    
};

g.getMailBodyHtml = function(mail) {

        if (!mail['Body']) {
            return '';
        }
        
        var _content = mail['Body']['Content'] || '', _contentType = mail['Body']['ContentType'];
        
        if (_contentType == 'HTML') {

// TODO: escape
            var $html = $('<div></div>').append(_content);
            $('script', $html).remove();
            _content = $html.html();
        } else {
            _content = $('<div></div>').text(_content).text().replace(/\r?\n/g, "<br />");
        }
        
        return _content;
};

g.getMailBodyPreview = function(mail) {

        if (!mail['BodyPreview']) {
            return '';
        }
        var _content = mail['BodyPreview'] || '';
        _content = $('<div></div>').text(_content).text().replace(/\r?\n/g, "<br />");
        
        return _content;
};

g.truncate = function(str,len){
    var _str = '' + str;
    if (_str.length <= len) {
        return _str;
    }
    return _str.substring(0, len) + '...';
};

})((this || 0).self || global);

$(function(){
});

