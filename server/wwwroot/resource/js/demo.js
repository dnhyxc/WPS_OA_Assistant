var isServerOk = false
var isSetupOk = false
var pluginsMode=0;//0表示jsplugins.xml模式，1表示publish模式，2表示 多进程加动态传递jsplugins.xml模式
//publish自动安装开始
 /*
    * 将自己的加载项地址配置到这里来
    * 需要保证加载项的name和业务业务系统中传递加载项name相对应
    * url必须以/ 结尾，且url+ribbon.xml和url+index.html在清除浏览器缓存的情况下能直接访问，不会被重定向
    * addonType:对应组件类型，wps文字，wpp演示，et表格
    */
    //复制开始
//    var curList = [{"name":"EtOAAssist","addonType":"et","online":"false","url":"http://127.0.0.1/jsplugindir/EtOAAssist.7z","version":"1.0.0"}]; //离线模式参考
var curList=[
    {"name":"WpsOAAssist","addonType":"wps","online":"true","url":"http://127.0.0.1:3888/plugin/wps/"},
    {"name":"EtOAAssist","addonType":"et","online":"true","url":"http://127.0.0.1:3888/plugin/et/"},
    {"name":"WppOAAssist","addonType":"wpp","online":"true","url":"http://127.0.0.1:3888/plugin/wpp/"}
]//在线模式配置参考
var localList=[];
var publishIndex=0;
var publishUnIndex=0;
/*获取用户本地全部加载项的接口是必须要的，这个接口做了判断，
** 如果58890端口未启动，会先去启动这个端口
*/
//加载项安装函数
function installWpsAddin(callBack){
    WpsAddonMgr.getAllConfig(function(e){
        
        
        if(!e.response||e.response.indexOf("null")>=0){//本地没有加载项，直接安装
            if(curList.length>0){
                installWpsAddinOne(callBack);
            } 
        }else{//本地有加载项，先卸载原有加载项，然后再安装
            localList=JSON.parse(e.response)
            // if(e.response)
            unInstallWpsAddinOne(callBack)
        }
            
    })
}
//安装单个加载项
function installWpsAddinOne(callBack){
    WpsAddonMgr.enable(curList[publishIndex],function(e){
        if(e.status){
            console.log(e.msg)
        }else{
            console.log("安装成功")
        }
        publishIndex++;
        if(publishIndex<curList.length){
            installWpsAddinOne();
        }else{
            callBack&&callBack()
        }
    })
}
//卸载单个加载项
function unInstallWpsAddinOne(callBack){
    WpsAddonMgr.disable(localList[publishUnIndex],function(e){
        if(e.status){
            console.log(e.msg)
        }else{
            console.log("卸载成功")
        }
        publishUnIndex++;
        if(publishUnIndex<localList.length){
            unInstallWpsAddinOne(callBack);
        }else{
            if(curList.length>0){
                installWpsAddinOne(callBack);
            } 
        }
    })
}
//publish自动安装结束
function envTest() {
    //1.服务端检测

    var xhr = getHttpObj()
    xhr.onload = function(e) {
        isServerOk = true
            //2.安装包检测
        var xhr1 = getHttpObj()
        xhr1.onload = function(e) {
            if (xhr1.responseText.indexOf('异常') > -1 || xhr1.responseText.indexOf('未安装') > -1 || xhr1.responseText.indexOf('失败') > -1) {
                window.location.href = "./demo.html"
                return;
            }
            isSetupOk = true
            window.location.href.indexOf("file://") > -1 ? window.location.href = "http://127.0.0.1:3888/index.html" : ''
        }
        xhr1.onerror = function(e) {
            isServerOk = false
            window.location.href = "./demo.html"
        }
        xhr1.open('get', 'http://127.0.0.1:3888/WpsSetupTest?pluginsMode='+pluginsMode, true)
        xhr1.send()
    }
    xhr.onerror = function(e) {
        isServerOk = false
        window.location.href = "./demo.html"
    }
    xhr.open('get', 'http://127.0.0.1:3888/FileList', true)
    xhr.send();
}
// 切换到相应 tab. 0: wps  1: wpp  2: et
function SwitchTab(crtTabIndex) {
    var iframe_wps = document.getElementById("iframe_wps");
    var iframe_wpp = document.getElementById("iframe_wpp");
    var iframe_et = document.getElementById("iframe_et");

    if (iframe_wps && iframe_wpp && iframe_et) {
        switch (crtTabIndex) {
            case 0:
                iframe_wpp.setAttribute("height", "0px");
                iframe_et.setAttribute("height", "0px");
                iframe_wps.setAttribute("height", "100%");
                break;
            case 1:
                iframe_wps.setAttribute("height", "0px");
                iframe_wpp.setAttribute("height", "0px");
                iframe_et.setAttribute("height", "100%");
                break;
            case 2:
                iframe_wps.setAttribute("height", "0px");
                iframe_et.setAttribute("height", "0px");
                iframe_wpp.setAttribute("height", "100%");
                break;
            default:
                iframe_wps.setAttribute("height", "0px");
                iframe_wpp.setAttribute("height", "0px");
                iframe_et.setAttribute("height", "0px");
                break;
        }
    }
}

function appChange() {
    var obj = document.getElementById("app");
    SwitchTab(obj.selectedIndex);
}

window.onload = function() {
    var app = document.getElementById("app");
    var opts = app.getElementsByTagName("option"); //得到数组option
    opts[0].selected = true;
    var indexFrame = document.getElementById("indexFrame");
    indexFrame.innerHTML="<iframe src='./resource/wps.html?pluginsMode="+pluginsMode+"' id='iframe_wps' frameborder='0' width='100%' height='0%'></iframe><iframe src='./resource/wpp.html?pluginsMode="+pluginsMode+"' id='iframe_wpp' frameborder='0' width='100%' height='0%'></iframe><iframe src='./resource/et.html?pluginsMode="+pluginsMode+"' id='iframe_et' frameborder='0' width='100%' height='0%'></iframe>"
    // 默认切换到 wps
    SwitchTab(0);
    //demo三种模式进行切换
    envTest()   //自动配置oem.ini
    if(pluginsMode==1){
        installWpsAddin() //自动安装集成
    }else{
        //卸载掉publish模式
        curList=[]
        installWpsAddin() 
    }
    
}

function IEVersion() {
    var userAgent = navigator.userAgent; //取得浏览器的userAgent字符串  
    var isIE = userAgent.indexOf("compatible") > -1 && userAgent.indexOf("MSIE") > -1; //判断是否IE<11浏览器  
    var isEdge = userAgent.indexOf("Edge") > -1 && !isIE; //判断是否IE的Edge浏览器  
    var isIE11 = userAgent.indexOf('Trident') > -1 && userAgent.indexOf("rv:11.0") > -1;
    if (isIE) {
        var reIE = new RegExp("MSIE (\\d+\\.\\d+);");
        reIE.test(userAgent);
        var fIEVersion = parseFloat(RegExp["$1"]);
        if (fIEVersion == 7) {
            return 7;
        } else if (fIEVersion == 8) {
            return 8;
        } else if (fIEVersion == 9) {
            return 9;
        } else if (fIEVersion == 10) {
            return 10;
        } else {
            return 6; //IE版本<=7
        }
    } else if (isEdge) {
        return 20; //edge
    } else if (isIE11) {
        return 11; //IE11  
    } else {
        return 30; //不是ie浏览器
    }
}

function getHttpObj() {
    var httpobj = null;
    if (IEVersion() < 10) {
        try {
            httpobj = new XDomainRequest();
        } catch (e1) {
            httpobj = new createXHR();
        }
    } else {
        httpobj = new createXHR();
    }
    return httpobj;
}
//兼容IE低版本的创建xmlhttprequest对象的方法
function createXHR() {
    if (typeof XMLHttpRequest != 'undefined') { //兼容高版本浏览器
        return new XMLHttpRequest();
    } else if (typeof ActiveXObject != 'undefined') { //IE6 采用 ActiveXObject， 兼容IE6
        var versions = [ //由于MSXML库有3个版本，因此都要考虑
            'MSXML2.XMLHttp.6.0',
            'MSXML2.XMLHttp.3.0',
            'MSXML2.XMLHttp'
        ];

        for (var i = 0; i < versions.length; i++) {
            try {
                return new ActiveXObject(versions[i]);
            } catch (e) {
                //跳过
            }
        }
    } else {
        throw new Error('您的浏览器不支持XHR对象');
    }
}