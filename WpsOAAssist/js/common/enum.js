// ====================== 源文件 ======================
/**
 * WPS常用的API枚举值，具体参与API文档
 */
var WPS_Enum = {
  wdDoNotSaveChanges: 0, // 不保存待定的更改
  // 更改：增加 wdPromptSaveChanges、wdSaveChanges
  wdPromptSaveChanges: -2, // 提示用户保存待定更改
  wdSaveChanges: -1, // 自动保存待定更改，而不提示用户。

  wdFormatPDF: 17,
  wdFormatOpenDocumentText: 23,
  wdFieldFormTextInput: 70,
  wdAlertsNone: 0,
  wdDialogFilePageSetup: 178,
  wdDialogFilePrint: 88,
  wdRelativeHorizontalPositionPage: 1,
  wdGoToPage: 1,
  wdPropertyPages: 14,
  wdRDIComments: 1,
  wdDialogInsertDateTime: 165,
  msoCTPDockPositionLeft: 0,
  msoCTPDockPositionRight: 2,
  /**
   * 将形状嵌入到文字中。
   */
  wdWrapInline: 7,
  /**
   * 将形状放在文字前面。 请参阅 wdWrapFront。
   */
  wdWrapNone: 3,
  /**
   * 使文字环绕形状。 行在形状的另一侧延续。
   */
  wdWrapSquare: 0,
  /**
   * 使文字环绕形状。
   */
  wdWrapThrough: 2,
  /**
   * 使文字紧密地环绕形状。
   */
  wdWrapTight: 1,
  /**
   * 将文字放在形状的上方和下方。
   */
  wdWrapTopBottom: 4,
  /**
   * 将形状放在文字后面。
   */
  wdWrapBehind: 5,
  /**
   * 将形状放在文字前面。
   */
  wdWrapFront: 6,
};

/**
 * WPS加载项自定义的枚举值
 */
var constStrEnum = {
  AllowOADocReOpen: "AllowOADocReOpen",
  AutoSaveToServerTime: "AutoSaveToServerTime",
  bkInsertFile: "bkInsertFile",
  // 更改：增加 bkInsertFileStart、bkInsertFileEnd
  /**
   * 套红书签开始
   */
  bkInsertFileStart: "bkInsertFileStart",
  /**
   * 套红书签结束
   */
  bkInsertFileEnd: "bkInsertFileEnd",

  buttonGroups: "buttonGroups",
  CanSaveAs: "CanSaveAs",
  copyUrl: "copyUrl",
  DefaultUploadFieldName: "DefaultUploadFieldName",
  disableBtns: "disableBtns",
  insertFileUrl: "insertFileUrl",
  IsInCurrOADocOpen: "IsInCurrOADocOpen",
  IsInCurrOADocSaveAs: "IsInCurrOADocSaveAs",
  isOA: "isOA",
  notifyUrl: "notifyUrl",
  OADocCanSaveAs: "OADocCanSaveAs",
  OADocLandMode: "OADocLandMode",
  OADocUserSave: "OADocUserSave",
  openType: "openType",
  picPath: "picPath",
  picHeight: "picHeight",
  picWidth: "picWidth",
  redFileElement: "redFileElement",
  revisionCtrl: "revisionCtrl",
  ShowOATabDocActive: "ShowOATabDocActive",
  SourcePath: "SourcePath",
  /**
   * 保存文档到业务系统服务端时，另存一份其他格式到服务端，其他格式支持：.pdf .ofd .uot .uof
   */
  suffix: "suffix",
  templateDataUrl: "templateDataUrl",
  TempTimerID: "TempTimerID",
  /**
   * 文档上传到业务系统的保存地址：服务端接收文件流的地址
   */
  uploadPath: "uploadPath",
  /**
   * 文档上传到服务端后的名称
   */
  uploadFieldName: "uploadFieldName",
  /**
   * 文档上传时的名称，默认取当前活动文档的名称
   */
  uploadFileName: "uploadFileName",
  uploadAppendPath: "uploadAppendPath",
  /**
   * 标志位： 1 在保存到业务系统时再保存一份suffix格式的文档， 需要和suffix参数配合使用
   */
  uploadWithAppendPath: "uploadWithAppendPath",
  userName: "userName",
  WPSInitUserName: "WPSInitUserName",
  taskpaneid: "taskpaneid",
  /**
   * 是否弹出上传前确认和成功后的确认信息：true|弹出，false|不弹出
   */
  Save2OAShowConfirm: "Save2OAShowConfirm",

  // 更改：以下所有变量为自定义添加的变量
  /**
   * 是否弹出关闭提示
   */
  CloseConfirmTip: "CloseConfirmTip",

  /**
   * 修订状态标志位
   */
  RevisionEnableFlag: "RevisionEnableFlag",

  // 用户自定义cookie
  CookieParams: "CookieParams",

  // 插入红头标识: 1：开始插入，2：插入成功
  InsertReding: "InsertReding",

  // 临时保存数据，用于业务保存pdf 和 nomarkPdf
  SaveAllTemp: "SaveAllTemp",

  // 企业ID
  OrgId: "orgId",

  // 新建文件名称
  NewFileName: "NewFileName",

  // 下载参数
  DownloadParams: "downloadParams",

  // 字段映射对象
  FieldObj: "fieldObj",

  // 正文字段 URL
  BodyTemplateUrl: "bodyTemplateUrl",

  // 需要禁用加载项按钮
  disabledBtns: "disabledBtns",
};

// 更改：增加套红字段类型
/**
 * 套红字段补充映射
 */
var fieldObjEnum = {
  标题: "title",

  文号: "refNo",

  /**
   * 外部来文文号
   */
  // OUTSIDE_REF_NO: 'outsideRefNo',

  // THEME_WORD: 'themeWord',

  /**
   * 公文类型
   */
  // OFFICIAL_TYPE: 'officialType',

  /**
   * 紧急程度
   */
  缓急: "urgencyLevel",

  /**
   * 密级
   */
  密级: "secretClass",

  // SECRECY_TERM: 'secrecyTerm',

  /**
   * 发送日期
   */
  // DISPATCH_DATE: 'dispatchDate',

  /**
   * 接收日期
   */
  // RECEIPT_DATE: 'receiptDate',

  /**
   * 接收单位
   */
  // DISPATCH_COMPANY: 'dispatchCompany',

  /**
   * 收文单位
   */
  // RECEIPT_COMPANY: 'receiptCompany',

  /**
   * 主送
   */
  主送: "mainSend",

  /**
   * 抄送
   */
  抄送: "copySend",

  签发人: "issUer",
  落款单位: "signingUnit",
  署名单位: "signatureUnit",
  签发日期: "issueDate",
  印发日期: "printDate",
  // PROOFREADER: 'proofreader',
  // PRINTER: 'printer',
  // IMPRESSION: 'impression',

  /**
   * 附件
   */
  附件: "enclosure",

  /**
   * 收文文号
   */
  // RECEIPT_REF_NO: 'receiptRefNo',

  /**
   * 行文类型
   */
  // WRITING: 'writing',

  /**
   * 是否用章
   */
  // USESEAL: 'useSeal',

  /**
   * 拟稿部门
   */
  // FILE_DEPARTMENT: 'fileDepartment',

  /**
   * 起草人
   */
  起草人: "creatPerson",
  /**
   * 部门
   */
  部门: "department",
  /**
   * 发文单位
   */
  发文单位: "units",

  /**
   * 创建时间/拟稿时间
   */
  // CREAT_TIME: 'creatTime',
};

// 更改：增加保存类型
/**
 * 保存类型
 */
var SAVE_TYPE = {
  /**
   * 无套红源文件
   */
  ORGINAL_NO_RED_HEAD: 1,

  /**
   * 无套红 PDF
   */
  PDF_NO_RED_HEAD: 2,

  /**
   * 套红源文件
   */
  ORGINAL_RED_HEAD: 3,

  // 套红html文件
  HTML_HEAD: 4,

  /**
   * 套红 PDF
   */
  PDF_RED_HEAD: 5,

  // 未套红html文件
  NO_HTML_HEAD: 6,
};
