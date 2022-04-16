var _et = {}

var pluginsMode = location.search.split("=")[1];//截取url中的参数值
var pluginType = WpsInvoke.ClientType.et//加载项类型wps,et,wpp
var pluginName = "EtOAAssist";//加载项名称
var wpsClient = new WpsClient(pluginType);//初始化一个多进程对象，多进程时才需要
var clientStr = pluginName + pluginType + "ClientId"
//单进程封装开始
/**
 * 此方法是根据wps_sdk.js做的调用方法封装
 * 可参照此定义
 * @param {*} funcs         这是在WPS加载项内部定义的方法，采用JSON格式（先方法名，再参数）
 * @param {*} front         控制着通过页面执行WPS加载项方法，WPS的界面是否在执行时在前台显示
 * @param {*} jsPluginsXml  指定一个新的WPS加载项配置文件的地址,动态传递jsplugins.xml模式，例如：http://127.0.0.1:3888/jsplugins.xml
 * @param {*} isSilent      隐藏打开WPS，如果需要隐藏，那么需要传递front参数为false
 */


function _WpsInvoke(funcs, front, jsPluginsXml,isSilent) {
    var info = {};
    info.funcs = funcs;
    if(isSilent){//隐藏启动时，front必须为false
        front=false;
    }
    /**
     * 下面函数为调起WPS，并且执行加载项WpsOAAssist中的函数dispatcher,该函数的参数为业务系统传递过去的info
     */
    if (pluginsMode != 2) {//单进程
        singleInvoke(info,front,jsPluginsXml,isSilent)
    } else {//多进程
        multInvoke(info,front,jsPluginsXml,isSilent)
    }

}

//单进程
function singleInvoke(info,front,jsPluginsXml,isSilent){
    WpsInvoke.InvokeAsHttp(pluginType, // 组件类型
        pluginName, // 插件名，与wps客户端加载的加载的插件名对应
        "dispatcher", // 插件方法入口，与wps客户端加载的加载的插件代码对应，详细见插件代码
        info, // 传递给插件的数据        
        function (result) { // 调用回调，status为0为成功，其他是错误
            if (result.status) {
                if (result.status == 100) {
                    WpsInvoke.AuthHttpesCert('请在稍后打开的网页中，点击"高级" => "继续前往"，完成授权。')
                    return;
                }
                alert(result.message)

            } else {
                console.log(result.response)
            }
        },
        front,
        jsPluginsXml,
        isSilent)

    /**
     * 接受WPS加载项发送的消息
     * 接收消息：WpsInvoke.RegWebNotify（type，name,callback）
     * WPS客户端返回消息： wps.OAAssist.WebNotify（message）
     * @param {*} type 加载项对应的插件类型
     * @param {*} name 加载项对应的名字
     * @param {func} callback 接收到WPS客户端的消息后的回调函数，参数为接受到的数据
     */
    WpsInvoke.RegWebNotify(pluginType, pluginName, handleOaMessage)
}
//多进程
function multInvoke(info,front,jsPluginsXml,isSilent){
    wpsClient.jsPluginsXml = jsPluginsXml ? jsPluginsXml : "https://127.0.0.1:3888/jsplugins.xml";
    if (localStorage.getItem(clientStr)) {
        wpsClient.clientId = localStorage.getItem(clientStr)
    }
    if(isSilent){
        wpsClient.StartWpsInSilentMode(pluginName,function(){//隐藏启动后的回调函数
            mult(info,front)
        })
    }else{
        mult(info,front)
    }
    wpsClient.onMessage = handleOaMessage
}
//多进程二次封装
function mult(info,front){
    wpsClient.InvokeAsHttp(
        pluginName, // 插件名，与wps客户端加载的加载的插件名对应
        "dispatcher", // 插件方法入口，与wps客户端加载的加载的插件代码对应，详细见插件代码
        info, // 传递给插件的数据        
        function (result) { // 调用回调，status为0为成功，其他是错误
            if (wpsClient.clientId) {
                localStorage.setItem(clientStr, wpsClient.clientId)
            }
            if (result.status !== 0) {
                console.log(result)
                if (result.message == '{\"data\": \"Failed to send message to WPS.\"}') {
                    wpsClient.IsClientRunning(function (status) {
                        console.log(status)
                        if (status.response == "Client is running.")
                            alert("任务发送失败，WPS 正在执行其他任务，请前往WPS完成当前任务")
                        else {
                            wpsClient.clientId = "";
                            wpsClient.notifyRegsitered = false;
                            localStorage.setItem(clientStr, "")
                            mult(info)
                        }
                    })
                    return;
                }
                else if (result.status == 100) {
                    // WpsInvoke.AuthHttpesCert('请在稍后打开的网页中，点击"高级" => "继续前往"，完成授权。')
                    return;
                }
                alert(result.message)
            } else {
                console.log(result.response)
            }
        },
        front)
}

function handleOaMessage(data) {
    console.log(data)
}
function GetUploadPath() {
    var url = document.location.host;
    return document.location.protocol + "//" + url + "/Upload";
}

function GetDemoPath(fileName) {

    var url = document.location.host;
    return document.location.protocol + "//" + url + "/file/" + fileName;
}

function newDoc() {
    _WpsInvoke([{
        "OpenDoc": {
            showButton: "btnSaveFile;btnSaveAsLocal"
        }
    }])
}

_et['newDoc'] = {
    action: newDoc,
    code: _WpsInvoke.toString() + "\n\n" + newDoc.toString(),
    detail: "\n\
  说明：\n\
    点击按钮，打开表格组件后,新建一个空白表格文档\n\
    \n\
  方法使用：\n\
    页面点击按钮，通过wps客户端协议来启动表格组件，调用oaassist插件，执行插件中的js函数OpenDoc,不带文档路径则默认新建一个空白表格文档\n\
    funcs参数说明:\n\
        OpenDoc方法对应于OA助手dispatcher支持的方法名\n\
            showButton 要显示的按钮\n\
"
}

function openDoc() {
    var filePath = prompt("请输入打开文件路径（本地或是url）：", GetDemoPath("样章.xlsx"))
    var uploadPath = prompt("请输入文档上传接口:", GetUploadPath())

    _WpsInvoke([{
        "OpenDoc": {
            "docId": "123", // 文档ID
            "uploadPath": uploadPath, // 保存文档上传接口
            "fileName": filePath,
            showButton: "btnSaveFile;btnSaveAsLocal"
        }
    }])
}

_et['openDoc'] = {
    action: openDoc,
    code: _WpsInvoke.toString() + "\n\n" + openDoc.toString(),
    detail: "\n\
  说明：\n\
    点击按钮，输入要打开的文档路径，输入文档上传接口，如果传的不是有效的服务端地址，将无法使用保存上传功能。\n\
    打开表格组件后,将根据文档路径下载并打开对应的文档，保存将自动上传指定服务器地址\n\
    \n\
  方法使用：\n\
    页面点击按钮，通过wps客户端协议来启动表格组件，调用oaassist插件，执行传输数据中的指令\n\
    funcs参数信息说明:\n\
        OpenDoc方法对应于OA助手dispatcher支持的方法名\n\
            docId 文档ID，OA助手用以标记文档的信息，以区分其他文档\n\
            uploadPath 保存文档上传接口\n\
            fileName 打开的文档路径\n\
            showButton 要显示的按钮\n\
"
}

function onlineEditDoc() {
    var filePath = prompt("请输入打开文件路径（本地或是url）：", GetDemoPath("样章.xlsx"))
    var uploadPath = prompt("请输入文档上传接口:", GetUploadPath())

    _WpsInvoke([{
        "OnlineEditDoc": {
            "docId": "123", // 文档ID
            "uploadPath": uploadPath, // 保存文档上传接口
            "fileName": filePath,
            showButton: "btnSaveFile"
        }
    }])
}

_et['onlineEditDoc'] = {
    action: onlineEditDoc,
    code: _WpsInvoke.toString() + "\n\n" + onlineEditDoc.toString(),
    detail: "\n\
  说明：\n\
    点击按钮，输入要打开的文档路径，输入文档上传接口，如果传的不是有效的服务端地址，将无法使用保存上传功能。\n\
    打开演示后,将根据文档路径下载并打开对应的文档，保存将自动上传指定服务器地址\n\
    \n\
  方法使用：\n\
    页面点击按钮，通过wps客户端协议来启动演示组件，调用oaassist插件，执行传输数据中的指令\n\
    funcs参数信息说明:\n\
        OnlineEditDoc方法对应于OA助手dispatcher支持的方法名\n\
            docId 文档ID，OA助手用以标记文档的信息，以区分其他文档\n\
            uploadPath 保存文档上传接口\n\
            fileName 打开的文档路径\n\
            showButton 要显示的按钮\n\
"
}

window.onload = function () {
    var btns = document.getElementsByClassName("btn");
    for (var i = 0; i < btns.length; i++) {
        btns[i].onclick = function (event) {
            document.getElementById("blockFunc").style.visibility = "visible";
            var btn2 = document.getElementById("demoBtn");
            btn2.innerText = this.innerText;
            document.getElementById("codeDes").innerText = _et[this.id].detail.toString();
            document.getElementById("code").innerText = _et[this.id].code.toString();
            document.getElementById("demoBtn").onclick = _et[this.id].action;
            hljs.highlightBlock(document.getElementById("code"));
        }
    }
}