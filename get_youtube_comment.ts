const MAX_RESULTS = 20;
const mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("main");
const database =  { 
  videoIdRange: () => SpreadsheetApp.getActiveSpreadsheet().getSheetByName("database").getRange('A2'),
  pageTokenRange: () => SpreadsheetApp.getActiveSpreadsheet().getSheetByName("database").getRange('B2'),
}
const ui = SpreadsheetApp.getUi();

/**
 * 親コメントの場合 name, comment, like, publishedAt, 空白
 * 子コメントの場合 '→', comment, like, publishedAt, name
 */
type OutputAttributes = [string, string, number, string, string][];

/** YoutubeApiからコメントを取得し、シートに出力するクラス */
class GetComments {
  private results: GoogleAppsScript.YouTube.Schema.CommentThreadListResponse;
  private outputAry: OutputAttributes;
  private i: number;
  private enabledDeepL: boolean

  constructor(enabledDeepL: boolean) {
    this.outputAry = [];
    this.i = 0;
    this.enabledDeepL = enabledDeepL;
  }

  /** YoutubeApiからコメントスレッド群を取得する */
  fetchCommentThreads() {
    this.results = YouTube.CommentThreads.list("snippet,replies", {
      videoId: database.videoIdRange().getValue(),
      maxResults: MAX_RESULTS,
      order: "relevance",
      pageToken: database.pageTokenRange().getValue(),
    });
  }

  /** コメントをシートに出力する */
  outputComments() {
    database.pageTokenRange().setValue(this.results.nextPageToken);

    this.results.items.forEach((item) => {
      this.pushTopLevelComment(item.snippet.topLevelComment.snippet);

      if (item.replies) {
        this.pushChildComments(item.replies.comments);
      }
    });

    mainSheet.getRange(mainSheet.getLastRow() + 1, 1, this.i, 5).setValues(this.outputAry);
  }

  /** 親コメントをoutputAryに代入する */
  private pushTopLevelComment = (
    topLevelSnippet: GoogleAppsScript.YouTube.Schema.CommentSnippet
  ) => {
    const { authorDisplayName, publishedAt, likeCount, textOriginal } =
      topLevelSnippet;
    const outputAry = this.outputAry;
    const i = this.i;

    // GoogleかDeepLで翻訳
    let translatedText: string
    if(this.enabledDeepL) {
      translatedText = translateUsingDeepL(textOriginal)
    } else {
      translatedText = LanguageApp.translate(textOriginal, '', 'ja')
    }

    outputAry.push([
      authorDisplayName,
      translatedText,
      likeCount,
      publishedAt,
      "",
    ]);
    
    this.i++;
  };

  /** 子コメントをoutputAryに代入する */
  private pushChildComments = (
    childComments: GoogleAppsScript.YouTube.Schema.Comment[]
  ) => {
    childComments.forEach((comment) => {
      const { authorDisplayName, publishedAt, likeCount, textOriginal } =
        comment.snippet;
      const outputAry = this.outputAry;
      const i = this.i;

      outputAry.push([
        "→",
        LanguageApp.translate(textOriginal, '', 'ja'),
        likeCount,
        publishedAt,
        authorDisplayName,
      ]);

      this.i++;
    });
  };
}

/** コメント取得とシート出力を実行するメソッド。引数にtrueを渡すとDeepLで翻訳 */
const fetchAndOutputComments = (enabledDeepL: boolean) => {
  let description = 'Youtube動画のURLを貼り付けてください。'
  // DeepLを使用する場合は現在の使用文字数と上限を取得する。この際apiキーがなければ例外が投げられる
  if(enabledDeepL) {
    description += '\n' + fetchDeepLQuota()
  }

  const res = ui.prompt(
    MAX_RESULTS + '件翻訳',
    description,
    ui.ButtonSet.OK_CANCEL
  );
  const url = res.getResponseText();

  if (res.getSelectedButton() !== ui.Button.OK || url === "") {
    ui.alert('キャンセルしました')
    return
  }

  // urlからvideoID抜き出し
  let videoId: string
  const regExp = /^.*((youtu.be\/)|(v\/)|(\/u\/\w\/)|(embed\/)|(watch\?))\??v?=?([^#\&\?]*).*/;
  const match = url.match(regExp);
  if (match&&match[7].length==11) { 
    videoId = match[7]
  } else {
    ui.alert('urlからvideoIdを抽出出来ませんでした。')
    return
  } 
    
  const getComments = new GetComments(enabledDeepL);
  try {
    database.videoIdRange().setValue(videoId)
    getComments.fetchCommentThreads();
    getComments.outputComments();
  } catch (e) {
    const isNotBeFound = e.message.endsWith('parameter could not be found.')
    const msg = isNotBeFound ? '検索がヒットしませんでした。。。' : e.message
    ui.alert(msg);
  }
};

/** 続きからコメント取得とシート出力。引数にtrueを渡すとDeepLで翻訳 */
const fetchAndOutputCommentsNext = (enabledDeepL: boolean) => {
  if(!database.videoIdRange().getValue()) {
    ui.alert('databaseシートのA2にvideoIdが存在しません')
    return
  }

  if(!database.pageTokenRange().getValue()) {
    ui.alert('databaseシートのB2にpageTokenが存在しません\n(続きのコメントは無いかもしれません)')
    return
  }


  let description = '前回取得したコメントの続きから実行します'
  if(enabledDeepL) {
    description += '\n' + fetchDeepLQuota()
  }
  const res = ui.alert(
    MAX_RESULTS + '件翻訳',
    description,
    ui.ButtonSet.OK_CANCEL
  );
  if (res !== ui.Button.OK) {
    ui.alert('キャンセルしました')
    return
  }

  const getComments = new GetComments(enabledDeepL);

  try {
    getComments.fetchCommentThreads();
    getComments.outputComments();
    ui.alert('出力が完了しました')
  } catch (e) {
    const isNotBeFound = e.message.endsWith('parameter could not be found.')
    const msg = isNotBeFound ? '検索がヒットしませんでした。。。' : e.message
    ui.alert(msg);
  }
};

/** 翻訳メソッドをGoogle翻訳かDeepLで実行 */
const fetchAndOutputCommentsGoogle = () => fetchAndOutputComments(false)
const fetchAndOutputCommentsNextGoogle = () => fetchAndOutputCommentsNext(false)
const fetchAndOutputCommentsDeepL = () => fetchAndOutputComments(true)
const fetchAndOutputCommentsNextDeepL = () => fetchAndOutputCommentsNext(true)


/** スプレッドシート読み込み時にYoutubeコメント翻訳メニューを追加する */
const onOpen = () => {
  // スプレッドシート全体を取得する
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet()

  // スプレッドシート整形
  formatSheet()
  
  const items = [
    { name: 'コメント' + MAX_RESULTS + '件翻訳(Google)', functionName: 'fetchAndOutputCommentsGoogle' },
    { name: '続けてコメント翻訳(Google)', functionName: 'fetchAndOutputCommentsNextGoogle' },
    null,
    { name: 'DeepL APIキーをセット', functionName: 'setDeepLApiKey' },
    { name: 'コメント' + MAX_RESULTS + '件翻訳(DeepL)', functionName: 'fetchAndOutputCommentsDeepL' },
    { name: '続けてコメント翻訳(DeepL)', functionName: 'fetchAndOutputCommentsNextDeepL' },
    null,
    { name: '初期化', functionName: 'clearSheet' },
  ]
  spreadSheet.addMenu('Youtubeコメント翻訳', items)
}

/** シート整形 */
const formatSheet = () => {
  const header = mainSheet.getRange('A1:E1')
  header.setValues([['名前(親)','コメント','いいね','日付','名前(子)']] )
  header.setBackground('#b7e1cd')
  mainSheet.setColumnWidth(2, 730)
  mainSheet.getRangeList(['B:B']).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
}

/** シート初期化 */
const clearSheet = () => {
  mainSheet.getRange(2,1, mainSheet.getMaxRows(), mainSheet.getMaxColumns()).clear()
  database.videoIdRange().clear()
  database.pageTokenRange().clear()
  formatSheet()
}
