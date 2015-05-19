(function(g){
"use strict";

g.Office365Api = {

	// AzureADを設定
	clientId:'12345678-1234-1234-1234-123456789012',
    callbackUrl:'http://localhost/callback',
    resourceUrl:'https://outlook.office365.com/',
    
    accessToken: null,
    refreshToken: null,
    me:null,
    mails:[],
    adMails:[],

    today: null,
    
    /**
     * 初期化及びTOP用データ読み込み 
     **/
    init:function() {
        
      var _this=this;
      var _dfd=$.Deferred();
      
      _this.refreshToken = localStorage.getItem('refresh_token') || null;
      _this.today = createToday();
      
      (function(){          
          if (!_this.refreshToken) {
            return _this.login.call(_this);
          } else {
            return _this.getAccessTokenByRefreshToken.call(_this);
          }
      })().then(function(){
            return _this.getUserInfo.call(_this);
          }
      ).then(function(){
          return _this.initMails.call(_this);
        }
      ).then(function(){
          console.log('finish');
          _dfd.resolve(_this);
        },
        function(e) {
          _dfd.reject(e);
        }
      );
      
      return _dfd.promise();
    },
    
    /**
     * SUBページ用初期化
     */
    initForSubPage:function(access_token) {
      var _this=this;
      
      _this.refreshToken = localStorage.getItem('refresh_token') || null;
      _this.now = createMyDate();
      _this.accessToken = access_token;
      
      return;
    },
    
    /**
     *  ログオフ（パラメータリセット）
     */
    logoff:function() {
        var _this=this;
        _this.accessToken=null;
        _this.refreshToken = null;
        _this.me = null;
        _this.mails = [];
        _this.adMails = [];
        _this.today = null;
        localStorage.removeItem('refresh_token');
    },
     

    
//    clearRefreshToken:function() {
//        var _this=this;
//        _this.refreshToken = null;
//        localStorage.removeItem('refresh_token');  
//    },

    /**
     * ログイン
     */
    login:function() {

        var _this = this;
        var _dfd = $.Deferred();

        // リソースタイプ、クライアントID、リダイレクトURL、リソース(end point api)
        var _url = 
            'https://login.microsoftonline.com/common/oauth2/authorize?'+
            'response_type=code'+
            '&client_id=' + _this.clientId + 
            '&redirect_uri=' + encodeURIComponent(_this.callbackUrl) + 
            '&resource=' + encodeURIComponent(_this.resourceUrl) ;        

        var _window = window.open(_url,'_blank','location=no,clearcache=yes');
        
        _window.addEventListener('loadstart', function(e) {
            
            // 一時コード取得
            if (e.url.indexOf(_this.callbackUrl)>=0) {
                var _code = $.url(e.url).param('code') || '';
                //$('body').append('code').append(_code);
                if (_code) {
                    // TOKEN取得
                    $.ajax({
                        type:'POST',
                        url: 'https://login.microsoftonline.com/common/oauth2/token',

                        dataType:'json',
                        cache: false,
                        data: {
                            // 他にclient_credentials, passwordがある
                            'grant_type':'authorization_code',
                            'code': _code,
                            'client_id':_this.clientId,
                            'redirect_uri': _this.callbackUrl
                        }
                    }).then(
                        function(data) {
                            _this.refreshToken = data.refresh_token || '';
                            localStorage.setItem( 'refresh_token', _this.refreshToken);
                            _this.accessToken = data.access_token || '';
                            _dfd.resolve();
                        },
                        function(e) {
                            _dfd.reject(e);
                        }
                    );                    
                }
                _window.close();
            }
        });
        return _dfd.promise();  
    },
    
    /**
     * アクセストークン取得 
     */
    getAccessTokenByRefreshToken:function() {
      var _this=this;
      var _dfd = $.Deferred();
      _this.accessToken = null;
      
      if (!_this.refreshToken) {
        _dfd.reject('error');
        return _dfd;
      }
        
        $.ajax({
            type: 'POST',
            url: 'https://login.microsoftonline.com/common/oauth2/token',
            data: {
                'grant_type':'refresh_token',
                'client_id' : _this.clientId,
                'refresh_token' : _this.refreshToken,
                'resorece' : _this.resourceUrl,
            },
        }).then(function(data){
                _this.accessToken = data.access_token || '';
                _dfd.resolve();
            },
            function(e) {
                _dfd.reject(e);
            }
        );
      
        return _dfd.promise();
    },
    
    /**
     * TOPページ用メール取得
     */
    initMails:function() {
        var _this=this, _dfd=$.Deferred();        
        var _mails=[], _filter = '';
        
        // 当日+既読
        _filter = 'DateTimeReceived ge ' + _this.today.toStringForApi() + ' and IsRead eq true';
        _this.getMails.call(_this, _filter).then(function(data){
            _mails = data;
//console.log('mail length:' + data.length)            ;            
            // 7日前+未読
            _filter = 'DateTimeReceived ge ' + _this.today.addTime(-1 * (7 * 24 * 3600 * 1000) ).toStringForApi() + ' and IsRead eq false';                    
            return _this.getMails.call(_this, _filter);
            
        }).then(function(data){
            
//console.log('data length:' + data.length)            ;            

        　　_mails = data.concat(_mails);
           _this.mails = [];
           _this.adMails = [];
//console.log('mail length:' + _mails.length)            ;
            var _ids=[];
            
           $.each(_mails, function(k,v) {
               var _id = v['Id'] || '';
               if (!_id) {
                   return;
               }
               // 処理済みの場合なし(重複している可能性もあるので）
               if (_ids.indexOf(_id)>=0) {
                   return;
               }
               _ids.push(_id);
               _this.checkMail.call(_this, v);
           });
           
            // 並べ替え
            _this.mails.sort(function(a,b){
                var a1=a['Priority'] || 9999, b1=b['Priority'] || 9999;
                if( a1 < b1 ) {
                    return -1;
                }                        
                if( a1 > b1 ) {
                    return 1;
                }                                            
                var a2 = a['DateTimeReceived'] || '', b2=b['DateTimeReceived'] || '';                    
                return (a2 > b2);
            });
            
           _dfd.resolve();
           
        }, function(e) {
           _dfd.reject(e); 
        });
        
        return _dfd;
    },
    
    getMails:function(filter) {
        
        var _this=this, _dfd=$.Deferred();
                        
        if (!_this.accessToken) {
            _dfd.reject('error');
            return _dfd.promise();
        }

        //var _date = _this.today.addTime(-1 * (7 * 24 * 3600 * 1000) ).toStringForApi();
        console.log('getMails ' + filter);

        $.ajax({
            type: 'GET',
            url:_this.resourceUrl + 'api/v1.0/me/folders/InBox/messages',
            data:{
                '$top':100,
                '$orderby':'DateTimeReceived desc',
                '$filter' : filter,
            },
            headers: {
                "Authorization": 'Bearer ' + _this.accessToken,
            },
            dataType: 'json',
            // async: false,
        }).then(
            function(data) {                
                _dfd.resolve((data['value'] || []));              
            },
            function(e) {
                console.log('fail');
                console.log(e);
                _dfd.reject(e);
            }
        );
        return _dfd.promise();
    },
    
    checkMail:function(mail) {
        
        var _this=this;
        var _flg;
        
        if (!mail) {
            return false;
        }

        if (mail['IsRead']) {
            var _date = _this.today.toYmd();
            if ((mail['DateTimeReceived'] || '').indexOf(_date) >= 0) {
                mail['Priority'] = 5;
                _this.mails.push(mail);
            }
            return true;
        }
        
        if (_this.checkAd.call(_this, mail)) {
            console.log('adMail push:' + (mail['Subject'] || ''));
            _this.adMails.push(mail);
            return true;
        }
        
        var _myaddr = _this.me['Id'] || '';
        
        var _to = mail['ToRecipients'] || [];
        var _cc = mail['CcRecipients'] || [];
        
        if (_to.length == 1 ) {
            // TO単数
            if (_to[0]['EmailAddress'] && _to[0]['EmailAddress']['Address'] && _to[0]['EmailAddress']['Address'] == _myaddr ) {
                mail['Priority'] = 1;
                _this.mails.push(mail);
                return true;
            }
        } else {
            // TO複数
            _flg = false;
            $.each(_to, function(k,v){
                if (v['EmailAddress'] && v['EmailAddress']['Address'] && v['EmailAddress']['Address'] == _myaddr ) {
                    _flg = true;
                    return false;
                }
            });
            if (_flg) {
                mail['Priority'] = 2;
                _this.mails.push(mail);
                return true;
            }
        }
        
        if (_cc.length > 0) {
            _flg = false;
            $.each(_cc, function(k,v){
                if (v['EmailAddress'] && v['EmailAddress']['Address'] && v['EmailAddress']['Address'] == _myaddr ) {
                    _flg = true;
                    return false;
                }
            });
            if (_flg) {
                mail['Priority'] = 3;
                _this.mails.push(mail);
                return true;
            }            
        }
        
        mail['Priority'] = 4;
        _this.mails.push(mail);
        
        return true;
    },
    
    checkAd:function(mail) {
        var _this=this;
        var _subject = mail['Subject'] || '';
        var _myaddr = _this.me['Id'] || '';
        
        console.log('checkAd:' + _subject + " " + _myaddr);
        
        var _flg = false;
        $.each(mail['ToRecipients'], function(k,v){
            if (v['EmailAddress'] && v['EmailAddress']['Address'] && v['EmailAddress']['Address'] == _myaddr) {
                _flg = true;
                return false;
            }
        });            
        $.each(mail['CcRecipients'], function(k,v){
            if (v['EmailAddress'] && v['EmailAddress']['Address'] && v['EmailAddress']['Address'] == _myaddr) {
                _flg = true;
                return false;
            }
        });            
        if (_flg) {
            return false;
        }
        
        var _from = '';
        if (mail['From'] && mail['From']['EmailAddress'] && mail['From']['EmailAddress']['Name']) {
            _from =  mail['From']['EmailAddress']['Name'];
        }
        if (!(_from.indexOf('事務局')>=0 || _from.indexOf('編集部')>=0)) {
             return false;
        }

        var _subject = mail['Subject'] || '';
        if (!(_subject.indexOf('企業様')>=0 || _subject.indexOf('無料')>=0 
        || _subject.indexOf('セミナー')>=0 || _subject.indexOf('お知らせ')>=0)) {
            return false;
        }

        var _body = (mail['BodyPreview'] || '').substring(0,80);
        if (!(_body.indexOf('本メール')>=0 || _body.indexOf('この電子メール')>=0 || _body.indexOf('本電子メール')>=0)) {
            return false;
        }
        
        return true;
    },
    
    /**
     * 関連メール取得
     */
    getMailsByConversationId:function(conversationId) {
        
        var _this=this;
        var _dfd=$.Deferred();
                        
        if (!_this.accessToken || !conversationId) {
            _dfd.reject('error');
            return _dfd.promise();
        }
        $.ajax({
            type: 'GET',
            url:_this.resourceUrl + 'api/v1.0/me/messages',
            data:{
                '$top':100,
                '$filter' :'ConversationId eq \'' + conversationId + '\'',
            },
            headers: {
                "Authorization": 'Bearer ' + _this.accessToken,
            },
            dataType: 'json',
        }).then(
            function(data) {
                _dfd.resolve(data['value']);              
            },
            function(e) {
                console.log('fail');
                console.log(e);
                _dfd.reject(e);
            }
        );
        return _dfd.promise();
    },

    setRead:function(id) {
        
        var _this=this;
        var _dfd=$.Deferred();
                        
        if (!_this.accessToken || !id) {
            _dfd.reject('error');
            return _dfd.promise();
        }

        $.ajax({
            type: 'PATCH',
            url:_this.resourceUrl + 'api/v1.0/me/messages/' + id,
            data: JSON.stringify({'IsRead':true}),
            headers: {
                "Authorization": 'Bearer ' + _this.accessToken,
            },
            contentType: 'application/json',
            processData: false,
            dataType: 'json',

        }).then(
            function(data) {
                console.log(data);
                _dfd.resolve(data);              
            },
            function(e) {
                console.log('fail');
                console.log(e);
                _dfd.reject(e);
            }
        );
        return _dfd.promise();
    },

    /**
     * 添付画像設定
     */ 
    setAttach:function($target, id) {
        
      var _this=this;
      var _dfd = $.Deferred();
      
      $target.html('');

      $.ajax({
            type: 'GET',
            url:_this.resourceUrl + 'api/v1.0/me/messages/'+id+'/attachments',
            headers: {
                "Authorization": 'Bearer ' + _this.accessToken,
            },
            dataType: 'json',
            // async: false,
        }).then(
            function(data) {
                
                var _res = null;
                if (data['value']) {                    
                    $.each(data['value'], function(k,v){
                       var _contentType = v['ContentType'] ||'';
                       if (_contentType.indexOf('image/jpeg')>=0 
                       || _contentType.indexOf('image/gif')>=0 
                       || _contentType.indexOf('image/png')>=0 ) {
                           _res = v;
                           return false;
                       }
                    });
                }
                
                // 画像設定
                if (_res !== null) {
                    var _bytes = _res['ContentBytes'] || '';
                    var _src = "data:" + _res['ContentType'] + ';base64,' + _bytes;
                    var $img = $('<img>', { 'src': _src });                    
                    $target.append($img);
                    _dfd.resolve();
                    return;
                }

                // 添付
                // 動作せず
/*
                if (data['value'] && data['value'][0]) {
                    _res = data['value'][0];
                    (function (res) {
                        var _bytes = atob(res['ContentBytes']);
                        var _type = res['ContentType'];
                        var _blob = new Blob([_bytes], { 'type': _type });
                        var _url = URL.createObjectURL(_blob);
                        var $a = $('<a>', { href: _url }).text('ダウンロード');
                        $target.append($a);
                    })(_res);
                    _dfd.resolve();
                    return;
                }
*/                
                    
                $target.hide();
                _dfd.resolve();
                return;
                
            },
            function(e) {
                console.log('fail');
                console.log(e);
                _dfd.reject(e);
            }
        );
        return _dfd.promise();
    },
    
    /**
     * ユーザ情報
     */
    getUserInfo:function() {

        var _this=this;
        var _dfd=$.Deferred();
        
        console.log('refreshUserInfo start');
        
        if (!_this.accessToken) {
            _dfd.reject();
            return _dfd.promise();
        }
        
        $.ajax({
            type: 'GET',
            url:_this.resourceUrl + 'api/v1.0/me',
            headers: {
                "Authorization": 'Bearer ' + _this.accessToken,
            },
            dataType: 'json',
        }).then(
            function(data) {
                _this.me = data;
                _dfd.resolve();
            },
            function(e) {
                console.log('fail');
                console.log(e);
                _dfd.reject(e);
            }
        );
        return _dfd.promise();
    },

    /**
     * プロフィール画像
     */
    getUserPhotoMe:function(){
        
        var _this=this;
        var _dfd=$.Deferred();

        if (!_this.accessToken || !_this.me || !_this.me['Id']) {
            _dfd.reject();
            return _dfd.promise();
        }

        var _imgUrl = _this.resourceUrl + 'ews/Exchange.asmx/s/GetUserPhoto?email='+_this.me['Id']+'&size=HR360x360';

        var _xhr = new XMLHttpRequest();
        _xhr.open('GET', _imgUrl, true);
        _xhr.setRequestHeader("Authorization", 'Bearer ' + _this.accessToken) ;
        _xhr.responseType = "arraybuffer";

        _xhr.onload = function() {
            
            var _bytes = new Uint8Array(this.response);
            var _raw = String.fromCharCode.apply(null,_bytes);


            var _imgSrc = '';
            if (_bytes[0] === 0xff && _bytes[1] === 0xd8 && _bytes[_bytes.byteLength-2] === 0xff && _bytes[_bytes.byteLength-1] === 0xd9) {
                _imgSrc = "data:image/jpeg;base64,";
            } else if (_bytes[0] === 0x89 && _bytes[1] === 0x50 && _bytes[2] === 0x4e && _bytes[3] === 0x47) {
                _imgSrc = "data:image/png;base64,";
            } else if (_bytes[0] === 0x47 && _bytes[1] === 0x49 && _bytes[2] === 0x46 && _bytes[3] === 0x38) {
                _imgSrc = "data:image/gif;base64,";
            } else if (_bytes[0] === 0x42 && _bytes[1] === 0x4d) {
                _imgSrc = "data:image/bmp;base64,";
            } else {
                _imgSrc = 'data:image/png;base64,';
            }
            
            _imgSrc += btoa(_raw);

            _dfd.resolve(_imgSrc);
        };
        _xhr.send();

        return _dfd.promise();
    },

    /**
     * 本日のイベント
     */
    getTodayEvents:function() {
        var _this=this;
        var _dfd=$.Deferred();
        
        if (!_this.accessToken) {
            _dfd.reject();
            return _dfd.promise();
        }
        var _startDateTime = _this.today.toStringForApi();
        var _endDateTime = _this.today.addTime(24*3600*1000 - 1).toStringForApi();
        

        $.ajax({
            type: 'GET',
            url:_this.resourceUrl + 'api/v1.0/me/calendarview',
            data: {
              'startDateTime':_startDateTime,
              'endDateTime':_endDateTime,
            },
            headers: {
                "Authorization": 'Bearer ' + _this.accessToken,
            },
            dataType: 'json',
        }).then(
            function(data) {
                _dfd.resolve(data);
            },
            function(e) {
                console.log('fail');
                console.log(e);
                _dfd.reject(e);
            }
        );
        
        return _dfd.promise();
    },
    
    /**
     * 本日のタスク(EWS利用)
     */
    getTasks:function() {
       var _this=this;
        var _dfd=$.Deferred();

        if (!_this.accessToken) {
            _dfd.reject('error');
            return _dfd.promise();
        }

        var _url = _this.resourceUrl + 'EWS/Exchange.asmx';

        var _data = '<?xml version="1.0" encoding="utf-8"?>' + "\n" +
'<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
' xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"' +
' xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">' + 
'<soap:Header>' +
'<t:RequestServerVersion Version="Exchange2013" />' +
//  '<t:TimeZoneContext>' +
//  '<t:TimeZoneDefinition Id="Tokyo Standard Time"/>' +
//  '</t:TimeZoneContext>' +
'</soap:Header>' +
'<soap:Body>' +
'<m:FindItem Traversal="Shallow">' +
'<m:ItemShape>' +
'<t:BaseShape>IdOnly</t:BaseShape>' +
'<t:AdditionalProperties>' +
          '<t:FieldURI FieldURI="item:Subject" />' +
          '<t:FieldURI FieldURI="task:DueDate" />' +
          '<t:FieldURI FieldURI="task:Status" />' +
          '<t:FieldURI FieldURI="task:IsComplete" />' +
'</t:AdditionalProperties>' +          
'</m:ItemShape>' +
'<m:ParentFolderIds>' +
'<t:DistinguishedFolderId Id="tasks"/>' +
'</m:ParentFolderIds>' +
//'<m:QueryString>task:IsComplete:false</m:QueryString>' +
'</m:FindItem>' +
'</soap:Body>' +
'</soap:Envelope>';

        $.ajax({
            type: 'POST',
            url: _url,
            data: _data,
            headers: {
                "Authorization": 'Bearer ' + _this.accessToken,
            },
            dataType: 'text',
        }).then(
            function(data) {
                // xml解析がどうもうまくいかないので。
                var _tasks=[];
                var _m = data.match(/<t\:Task>.+?<\/t\:Task>/igm);
                if (_m.length ) {
                    $.each(_m, function(k,v){
                        var _item = {};
                        _item['Subject'] = '';
                        if (v.match(/<t\:Subject>(.*?)<\/t:Subject>/im)) {
                            _item['Subject'] = RegExp.$1;
                        }
                        _item['DueDate'] = null;
                        _item['DueDate2'] = createMyDate('1900-01-01'); // 初期値
                        if (v.match(/<t\:DueDate>(.*?)<\/t:DueDate>/im)) {
                            _item['DueDate'] = RegExp.$1;
                            _item['DueDate2'] = createMyDate(RegExp.$1);
                        }
                        _item['Status'] = null;
                        if (v.match(/<t\:Status>(.*?)<\/t:Status>/im)) {
                            _item['Status'] = RegExp.$1;
                        }
                        if (_item['Status'].toLowerCase() == 'completed') {
                            return;
                        }
                        _tasks.push(_item);
                    });
                }
                
                // 今日に近いもの
                var _nowTime = _this.today.getTime();
                _tasks.sort(function(a,b){
                    var a1 = Math.abs(a['DueDate2'].getTime() - _nowTime),
                        b1 = Math.abs(b['DueDate2'].getTime() - _nowTime);
                        if (a1 > b1) {
                            return 1;
                        }
                        return -1;
                });
                
                _dfd.resolve(_tasks);              
            },
            function(e) {
                console.log('fail');
                console.log(e);
                _dfd.reject(e);
            }
        );
        return _dfd.promise();        
    },
};

// common function
//g.getOffice365Api = function(){    
//  console.log('office365Api');
//  return Office365Api;  
//};

})((this || 0).self || global);

