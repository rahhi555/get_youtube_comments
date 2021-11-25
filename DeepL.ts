/** DeepLApiキーをプロパティサービスでセット及びゲットするときの文字列 */
const DEEPL_API_KEY = 'DEEPL_API_KEY'

/** DeepLAPIキーをプロパティサービス(GASの環境変数的なもの)にセットするメソッド */
const setDeepLApiKey = () => {
  const apiKey = PropertiesService.getUserProperties().getProperty(DEEPL_API_KEY)
  const msg = apiKey ? `現在のキーは${apiKey}です。\n上書きするならOKを、操作を取り消すならキャンセルを押してください\n` + 
                       '(APIキーを削除する場合は空白のままOKを押してください)'  : 'DeepL APIキーは未登録です'
  const res = ui.prompt(msg, ui.ButtonSet.OK_CANCEL)
  if(res.getSelectedButton() !== ui.Button.OK) {
    ui.alert('キャンセルしました')
    return
  }

  PropertiesService.getUserProperties().setProperty(DEEPL_API_KEY, res.getResponseText())
  ui.alert('DeepL APIキーをセットしました。 APIキー:' + res.getResponseText())
}

/** DeepLのご利用文字数と残り文字数を取得 */
const fetchDeepLQuota = () => {
  const apiKey = PropertiesService.getUserProperties().getProperty(DEEPL_API_KEY)
  if(!apiKey) throw new Error('DeepL APIキーがセットされていません。APIキーをセットしてから再度実行してください。')

  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: "get",
    headers: { 'Authorization': `DeepL-Auth-Key ${apiKey}` },
    muteHttpExceptions: true
  }

  const res = UrlFetchApp.fetch('https://api-free.deepl.com/v2/usage', options)
  const errorMsg = returnDeepLErrorMsg(res.getResponseCode())
  // errorMsgがnullでければエラーが発生したということなので、例外を投げてスクリプト全体の処理を終了させる
  if(errorMsg) throw new Error(errorMsg)

  const json = res.getContentText()
  const { character_count, character_limit } = JSON.parse(json) as { character_count: number, character_limit: number }
  return `(DeepL使用状況 : ${character_count} / ${character_limit} )`
}

/** DeepLによる翻訳 */
const translateUsingDeepL = (text: string)  => {
  const apiKey = PropertiesService.getUserProperties().getProperty(DEEPL_API_KEY)

  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: "post",
    headers: { 'Authorization': `DeepL-Auth-Key ${apiKey}` },
    payload: `text=${text}&target_lang=JA`,
    muteHttpExceptions: true
  }

  const res = UrlFetchApp.fetch('https://api-free.deepl.com/v2/translate', options)
  const errorMsg = returnDeepLErrorMsg(res.getResponseCode())

  if(errorMsg) throw new Error(errorMsg)
  
  const json = res.getContentText()
  const data = JSON.parse(json)

  return data.translations[0].text as string
}

/** レスポンスコードがエラーコードに該当していればエラーメッセージを返し、そうでなければnullを返す */
const returnDeepLErrorMsg = (responseCode: number) => {
  switch(responseCode) {
    case 400:
      return '不正なリクエストです。エラーメッセージとパラメータを確認してください。'
      break;
    case 403:
      return '認証に失敗しました。有効なDeepL APIキーを指定してください。'
      break;
    case 404:
      return 'リクエストされたリソースが見つかりませんでした。'
      break;
    case 413:
      return 'リクエストサイズが制限値を超えています。'
      break;
    case 414:
      return 'リクエストURLが長すぎます。GETリクエストではなくPOSTリクエストを使用し、パラメータをHTTPボディで送信することで、このエラーを回避することができます。\n' +
             '(POSTメソッドでリクエストをしているので、このエラーメッセージが現れることはない...はず)'
      break;
    case 429:
      return 'リクエストが多すぎます。しばらく待ってから、再度リクエストを送信してください。'
      break;
    case 456:
      return 'クォータ(割り当て量)を超えました。文字数制限に達しました。'
      break;
    case 503:
      return 'リソースは現在利用できません。後でもう一度お試しください。'
      break
    case 529:
      return 'リクエストが多すぎます。しばらく待ってから、再度リクエストを送信してください。'
      break
    default:
      return null
      break
  }
}