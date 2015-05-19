(function(g){
"use strict";

g.TopPage = {

    o: null,

    init:function() {

      console.log('top');

      var _this = this;
      _this.o = Office365Api;

      _this.o.init.call(_this.o).then(function(){
          _this.display.call(_this);
      });

      return;
    },

    display:function(){

        console.log('displayMain');

        var _this=this;
        var _dfd=$.Deferred();

        $.ajax({
            type:'GET',
            url:'top.html',
            dataType:'text',
        }).then(function(data){

            var $html=$(data).filter('div#wrapper');
            // <div>2015.4.3 Fri<br><span class="logo">パノラマ新聞</span></div>
            var _titleDate = createMyDate();
            $('div#header-title', $html).html('<div>' + _titleDate.toDisplayTitle() + '<br><span class="logo">パノラマ新聞</span></div>');

            // プロフィール写真
            _this.o.getUserPhotoMe.call(_this.o).then(function(data){
                $('#header-user-pic img', $html)
                .css({
                    "-webkit-filter": "grayscale(1)",
                    "filter": "grayscale(1)"
                }).attr({
                   src: data, width: '105px',
                });
            });

            // スケジュール
            $('dl#schedule dd', $html).empty();
            _this.o.getTodayEvents.call(_this.o).then(function(data){
               if (data['value']) {
                   $('dl#schedule dd', $html).each(function(i){
                      if (data['value'][i]) {
                          $(this).text(truncate(data['value'][i]['Subject'] || '', 25));
                      }
                   });
               }
            });

            // タスク
            $('dl#task dd', $html).text('');
            _this.o.getTasks.call(_this.o).then(function(data) {
    			$('#task dd', $html).each(function(i){
                    if (data[i]) {
                        $(this).text(truncate(data[i]['Subject'] || '', 25));
                    }
                });
            });

            // 広告クリック
            $('a.more-ad', $html).off('click').on('click',function(e){
                localStorage.setItem('ad_mails', JSON.stringify(_this.o.adMails || []));
                var _window = window.open('ad_list.html','_blank','location=no,clearcache=yes');
                return false;
            });

            // 広告
            $('dl#ad dd', $html).text('');
            $('dl#ad dd', $html).each(function(k){
                var _subject = (_this.o.adMails[k] && _this.o.adMails[k]['Subject']) || '';
                //console.log("Subject:" + _subject);
                $(this).text(truncate(_subject, 25));
            });

            // メール設定
            for (var i=1;i<=20;i++){
              var _class = '.mail-blk-' + (i<10 ? '0' : '') + i;
              var $inner = $(_class, $html);
                if ($inner.length == 0) {

                    continue;
                }

                var _mail = _this.o.mails[i - 1] || null;

                if (_mail === null) {
                    $inner.html('');
                    continue;
                }

                


                $('.mail-title', $inner).text(_mail['Subject'] || '');
                //$('.mail-title', $inner).append(':'+_mail['Priority']);

                var _dt = createMyDate(_mail["DateTimeReceived"]);

                $('.mail-date', $inner).text(_dt.toDisplay());

                var _content = getMailBodyPreview(_mail);
                $('.mail-text', $inner).html(_content);

                var _id = _mail['Id'] || '';

                if (_mail['HasAttachments']) {
                  $('li.btn-attach', $inner).addClass('on').text('添付あり');
                  _this.o.setAttach.call(_this.o, $('.mail-img', $inner), _id);
                } else {
                    $('.mail-img', $inner).hide();                    
                  $('li.btn-attach', $inner).removeClass('on').text('添付なし');

                }

                var $read = $('.btn-read', $inner)
                if (_mail['IsRead']){
                  $read.addClass('on');
                  $('a', $read).text('既読');
                } else {
                  $read.removeClass('on');
                  $('a', $read).text('未読');
                }

                (function(mail){
                  // もっと見る
                    $('.more', $inner).off('click').on('click',function(e){
                        localStorage.setItem('mail_detail', JSON.stringify(mail));
                        var _window = window.open('mail_detail.html','_blank','location=no,clearcache=yes');
                        return false;
                    });

                    // 同じスレッドのメールを見る
                    $('.btn-sled a', $inner).off('click').on('click', function(e){
                        localStorage.setItem('mail_detail', JSON.stringify(mail));
                        localStorage.setItem('access_token', _this.o.accessToken);
                        var _window = window.open('mail_list.html','_blank','location=no,clearcache=yes');
                        return false;
                    });

                    // 既読にする
                    $('.btn-read a', $inner).off('click').on('click', function(e){
                      var $li = $(this).closest('li');
                      if ($li.is('.on')) {
                        return false;
                      }
                      $li.addClass('on');
                      $(this).text('既読');
                      _this.o.setRead.call(_this.o, (mail['Id'] || ''));
                      return false;
                    });

                })(_mail);
            }

            $('body').empty().append($html);

        });

        return;
    },

};

})((this || 0).self || global);
