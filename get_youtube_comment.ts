const MAX_RESULTS = 25;
const mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("main");
const database =  { 
  videoIdRange: () => SpreadsheetApp.getActiveSpreadsheet().getSheetByName("database").getRange('A2'),
  pageTokenRange: () => SpreadsheetApp.getActiveSpreadsheet().getSheetByName("database").getRange('B2'),
}

/**
 * 親コメントの場合 name, comment, like, publishedAt, 空白
 * 子コメントの場合 空白, comment, like, publishedAt, name
 *
 * 子コメントは親コメントに対して一列右にずれているような見た目にする
 */
type OutputAttributes = [string, string, number, string, string][];

/** YoutubeApiからコメントを取得し、シートに出力するクラス */
class GetComments {
  private results: GoogleAppsScript.YouTube.Schema.CommentThreadListResponse;
  private outputAry: OutputAttributes;
  private i: number;

  constructor() {
    this.outputAry = [];
    this.i = 0;
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

    outputAry.push([
      authorDisplayName,
      LanguageApp.translate(textOriginal, '', 'ja'),
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

/** コメント取得とシート出力を実行するメソッド */
const fetchAndOutputComments = () => {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(
    MAX_RESULTS + '件翻訳',
    "Youtube動画IDを入力(URL最後の文字列)",
    ui.ButtonSet.OK_CANCEL
  );
  const text = res.getResponseText();

  if (res.getSelectedButton() !== ui.Button.OK || text === "") return;

  const getComments = new GetComments();
  try {
    database.videoIdRange().setValue(text)
    getComments.fetchCommentThreads();
    getComments.outputComments();
  } catch (e) {
    const isNotBeFound = e.message.endsWith('parameter could not be found.')
    const msg = isNotBeFound ? '検索がヒットしませんでした。。。' : e.message
    Browser.msgBox(msg);
  }
};

/** 続きからコメント取得とシート出力 */
const fetchAndOutputCommentsNext = () => {
  if(!database.videoIdRange().getValue()) {
    Browser.msgBox('databaseシートのA2にvideoIdが存在しません')
    return
  }

  if(!database.pageTokenRange().getValue()) {
    Browser.msgBox('databaseシートのB2にpageTokenが存在しません')
    return
  }

  const getComments = new GetComments();
  try {
    getComments.fetchCommentThreads();
    getComments.outputComments();
  } catch (e) {
    const isNotBeFound = e.message.endsWith('parameter could not be found.')
    const msg = isNotBeFound ? '検索がヒットしませんでした。。。' : e.message
    Browser.msgBox(msg);
  }
};

/** シート初期化 */
const clearSheet = () => {
  mainSheet.getRange(2,1, mainSheet.getMaxRows(), mainSheet.getMaxColumns()).clear()
  database.videoIdRange().clear()
  database.pageTokenRange().clear()
}

/** スプレッドシート読み込み時にYoutubeコメント翻訳メニューを追加する */
const onOpen = () => {
  // スプレッドシート全体を取得する
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet()
  const items = [
    { name: 'コメント' + MAX_RESULTS + '件翻訳', functionName: 'fetchAndOutputComments' },
    { name: '続けてコメント翻訳', functionName: 'fetchAndOutputCommentsNext' },
    null,
    { name: '初期化', functionName: 'clearSheet' },
  ]
  spreadSheet.addMenu('Youtubeコメント翻訳', items)
}