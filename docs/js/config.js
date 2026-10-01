/**
 * アプリの設定。GitHub Pages に公開する前に clientId と shareUrl を入れる。
 *
 * どちらも秘密情報ではない(公開してよい):
 *   - clientId : Entra のアプリの「アプリケーション (クライアント) ID」
 *   - shareUrl : OneDrive の Excel の共有リンク。共有されたアカウントでサインインした人しか開けない
 * 空のままにすると、画面で入力した値(そのブラウザにだけ保存)を使う。
 */
window.APP_CONFIG = {
  clientId: "10997d16-6ad4-4622-8ede-004a34f06027",
  shareUrl: "https://1drv.ms/x/c/9E6F1B8A9E5D1375/AdFtwOfLD7hDgtHuw1FuObk?e=cO1q4Z",

  // Microsoft Graph に要求する権限。フェーズ2の試作で「Files.ReadWrite も要求する」に
  // チェックしないと読めなかった場合は "Files.ReadWrite" を足す(Entra 側にも追加が必要)。
  scopes: ["User.Read", "Files.Read.All"],

  // スケジュール画面が今日から何週間先・何週間前まで表示するか
  weeksAhead: 10,
  weeksBehind: 1,
};
