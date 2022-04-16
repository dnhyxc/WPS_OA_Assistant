// ====================== 源文件 ======================

/**
 * 在这个js中，集中处理来自OA的传入参数
 *
 */

/**
 * web页面调用WPS的方法入口
 *  * info参数结构
 * info:[
 *      {
 *       '方法名':'方法参数',需要执行的方法
 *     },
 *     ...
 *   ]
 * @param {*} info
 */
let testFuncs = null;
function dispatcher(info) {
  console.log("info", info);
  var funcs = info.funcs;
  wps.WpsApplication().Options.UseLocalUserInfo = true; //修订时，不使用登录名

  // 更改：设置必要的 cookie
  InitCustomCookie(info.cookieParams);

  //解析参数
  /**
   * 分两种情况解析：
   * 1、业务系统依赖客户端返回：做同步处理，直接在for循环中返回
   */
  for (var index = 0; index < funcs.length; index++) {
    var func = funcs[index];
    for (var key in func) {
      if (key === "GetDocStatus") {
        return GetDocStatus(func[key]);
      } else if (key === "ExitWPS") {
        ExitWPS(func[key]);
      }
    }
  }
  /**
   * 2、业务系统不依赖客户端返回：
   * 做异步处理，先返回接收状态，再在setTimeout中做打开文档的一系列业务逻辑
   */
  setTimeout(function () {
    for (var index = 0; index < funcs.length; index++) {
      testFuncs = funcs;
      var func = funcs[index];
      for (var key in func) {
        func[key].isOA = true;
        if (key === "OpenDoc") {
          // OpenDoc 属于普通的打开文档的操作方式，文档落地操作
          OpenDoc(func[key]); //进入打开文档处理函数
        } else if (key === "OnlineEditDoc") {
          //在线方式打开文档，属于文档不落地的方式打开
          OnlineEditDoc(func[key]);
        } else if (key === "NewDoc") {
          OpenDoc(func[key]);
        } else if (key === "UseTemplate") {
          OpenDoc(func[key]);
        } else if (key === "InsertRedHead") {
          InsertRedHead(func[key]);
        } else if (key === "taskPaneBookMark") {
          taskPaneBookMark(func[key]);
        } else if (key === "NewOfficialDocument") {
          return OpenDoc(func[key]);
        }
      }
    }
  }, 100);
  return { message: "ok", app: wps.WpsApplication().Name };
}

/**
 * 获取活动文档的状态
 */
// function GetDocStatus() {
//   let l_doc = wps.WpsApplication().ActiveDocument;
//   if (l_doc && pCheckIfOADoc()) {
//     //此方法还可根据需要进行扩展
//     return {
//       message: "GetDocStatus",
//       docstatus: {
//         words: l_doc.Words.Count,
//         saved: l_doc.Saved,
//         pages: l_doc.ActiveWindow.Panes.Item(1).Pages.Count,
//       },
//     };
//   }
// }

// 更改：改变message中的参数
function GetDocStatus() {
  var wpsApp = wps.WpsApplication();
  let l_doc = wpsApp.ActiveDocument;
  if (l_doc && pCheckIfOADoc()) {
    //此方法还可根据需要进行扩展
    return {
      action: "status",
      message: {
        docNum: wpsApp.Documents.Count,
        docId: l_doc.DocID,
        words: l_doc.Words.Count,
        saved: l_doc.Saved,
        pages: l_doc.ActiveWindow.Panes.Item(1).Pages.Count,
      },

      //   message: "GetDocStatus",
      //   docstatus: {
      //     words: l_doc.Words.Count,
      //     saved: l_doc.Saved,
      //     pages: l_doc.ActiveWindow.Panes.Item(1).Pages.Count,
      //   },
    };
  }
  return {
    action: "status",
  };
}

/**
 * 关闭WPS活动文档并退出WPS进程
 */
// function ExitWPS() {
//   //为了接下来的用户确认操作，先将WPS界面激活显示在最前
//   wps.WpsApplication().Activate();
//   let l_doc = wps.WpsApplication().ActiveDocument;
//   if (l_doc && pCheckIfOADoc()) {
//     //此方法还可根据需要进行扩展
//     l_doc.Close();
//   }
//   if (
//     wps.confirm(
//       "要关闭WPS软件，请确认文档都已保存。\n点击确定后关闭WPS，点击取消继续编辑。"
//     )
//   ) {
//     wps.WpsApplication().Quit();
//   }
// }

// 更改：提示时根据传入的OaParams参数增加相应的判断
function ExitWPS(OaParams) {
  //为了接下来的用户确认操作，先将WPS界面激活显示在最前
  wps.WpsApplication().Activate();
  let l_doc = wps.WpsApplication().ActiveDocument;
  if (l_doc && pCheckIfOADoc()) {
    //此方法还可根据需要进行扩展
    l_doc.Close();
  }
  if (
    !OaParams.force &&
    wps.confirm(
      "要关闭WPS软件，请确认文档都已保存。\n点击确定后关闭WPS，点击取消继续编辑。"
    )
  ) {
    wps.WpsApplication().Quit();
  }
}

/**
 *
 * @param {*} params  OA端传入的参数
 */
function OnlineEditDoc(OaParams) {
  if (OaParams.fileName == "") {
    NewFile(OaParams);
  } else {
    //OA传来下载文件的URL地址，调用openFile 方法打开
    OpenOnLineFile(OaParams);
  }
}

// 更改：增加初始化cookie的方法
/**
 * 初始化cookie的方法
 */
// 初始化 cookie
function InitCustomCookie(cookieParams) {
  if (!cookieParams) {
    return;
  }
  Object.keys(cookieParams || {}).forEach((key) => {
    Cookies.set(key, cookieParams[key]);
  });
  // wps.PluginStorage.setItem(constStrEnum.CookieParams, params.cookieParams)
}

// 更改：增加初始化自定义参数的方法：InitCustomParam
// 初始化自定义参数
function InitCustomParam(OaParams) {
  // 插入红头标识
  if (OaParams.insertFileUrl) {
    wps.PluginStorage.setItem(constStrEnum.InsertReding, 1);
  }

  // 企业id
  // if (OaParams.orgId) {
  //   wps.PluginStorage.setItem(constStrEnum.OrgId, OaParams.orgId)
  // }
}

// 不需要的自定义方法
// function getOrgId() {
//   return wps.PluginStorage.getItem(constStrEnum.OrgId)
// }

// 打开来自OA端传递来的文档
// function OpenDoc(OaParams) {
//   if (OaParams.fileName == "") {
//     NewFile(OaParams);
//   } else {
//     //OA传来下载文件的URL地址，调用openFile 方法打开
//     OpenFile(OaParams);
//   }
// }

// 更改：首先进行初始化参数InitCustomParam(OaParams)
// 打开来自OA端传递来的文档
function OpenDoc(OaParams) {
  InitCustomParam(OaParams);
  if (OaParams.fileName == "") {
    NewFile(OaParams);
  } else {
    //OA传来下载文件的URL地址，调用openFile 方法打开
    OpenFile(OaParams);
  }
}

function taskPaneBookMark(OaParams) {
  let filePath = OaParams.fileName;
  if (filePath == "") return;
  OpenFile(OaParams);

  //创建taskpane，只创建一次
  let id = wps.PluginStorage.getItem(constStrEnum.taskpaneid);
  if (id) {
    let tp = wps.GetTaskPane(id);
    tp.Width = 300;
    tp.Visible = true;
  } else {
    let url = getHtmlURL("taskpane.html");
    let tp = wps.CreateTaskPane(url, "书签操作");
    if (tp) {
      tp.DockPosition = WPS_Enum.msoCTPDockPositionRight; //这里可以设置taskapne是在左边还是右边
      tp.Width = 300;
      tp.Visible = true;
      wps.PluginStorage.setItem(constStrEnum.taskpaneid, tp.ID);
    }
  }
}
