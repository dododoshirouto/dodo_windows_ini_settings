以下のようなwinget-list.txtを見て、wingetしていくシステムを作る

```
Git.Git
Google.Chrome
# GeekUninstaller.GeekUninstaller
```

`#`から始まる行は無視する（コメント）

IDで記述されていたらそれを、名前で記述されていたらそれを、インストールする
インストールはなるべくQuietに実効する
エラーが出たら、その行を飛ばして次へ進む
最後にエラーが出たものをリストにして表示する
