<!-- 設計メモ -->

# 品質とリスク

## リスクマトリクス

| リスク                                        | 影響   | 発生可能性 | ローンチ上の扱い             | 根拠                                                                                                                                                                                                       |
| --------------------------------------------- | ------ | ---------- | ---------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| go.mod と CI / release の Go バージョン不整合 | 低     | 低         | 対応済みの監視項目           | 現在の CI / release workflow は `actions/setup-go` の `go-version-file: go.mod` を使用しており、README / CONTRIBUTING にも `go.mod` を supported toolchain の source of truth として明記した。             |
| Excel 実機 E2E の不足                         | 高     | 中〜高     | 要対策                       | リポジトリ自身の hardening 文書が blank workbook / class / UserForm / .frx / init の repeatable verification を今後の完了条件としており、README / CONTRIBUTING も COM E2E は実機でやる前提です。2          |
| ドキュメント / 仕様ドリフト                   | 中     | 中         | 継続監視                     | ADR-0004、CLI 契約、README、CHANGELOG の session / version / scaffold update-check 周辺は現行実装に合わせて更新した。残る drift は release gate と license / vuln 運用の明文化が中心。                     |
| 依存脆弱性の継続監視不足                      | 中     | 中         | ローンチ前に最低限 scan 推奨 | セキュリティポリシー自体はありますが、公開設定から govulncheck、SBOM 生成、Go vulnerability DB 連携の自動化は確認できません。Go 公式は govulncheck と vuln.go.dev を推奨しています。14                     |
| プライバシー説明不足                          | 低〜中 | 低〜中     | 追跡継続                     | README / CONTRIBUTING に scaffold 時の GitHub Releases API 参照、`go install` 時の module mirror / checksum DB、`--no-update-check` / `XLFLOW_NO_UPDATE_CHECK` を追記した。独立した privacy 文書は未作成。 |
| ライセンス台帳の更新漏れ                      | 中     | 中         | ローンチ前修正が望ましい     | THIRD_PARTY_LICENCES.md の注記が現在の LICENSE 状態と噛み合っていません。go-localereader の manual review も残っています。12                                                                               |

## 必須修正と推奨改善

| 優先度 | 時間軸 | 施策                         | 具体的に直すべき点                                                                                                                                                       |
| ------ | ------ | ---------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| 最優先 | 短期   | 実機 E2E の release gate     | blank workbook / module / class / UserForm / .frx / init を含む Windows + Excel 実機確認を release checklist に組み込む。これは repo 自身の硬化文書が要求しています。2   |
| 高     | 短期   | ドキュメント整合の維持       | session auto-reuse / `save` / `version --verbose` / scaffold update-check opt-out を今後も CLI 契約、ADR、README、CHANGELOG の同一変更内で更新する。                     |
| 高     | 短期   | 公開上の期待値制御           | リリース名か README 先頭で「Preview」「Windows-only」「Excel COM / VBIDE trust required」「unsigned binary if applicable」を明示する。winget の有無よりここが重要です。3 |
| 高     | 短期   | プライバシー説明             | README / CONTRIBUTING へ追加した outbound-network 説明を独立した privacy / networking note に切り出すか、このまま運用するかを決める。                                    |
| 中     | 中期   | vuln / license scan の自動化 | CI に govulncheck と license scan を追加し、release 直前の手動運用から脱却する。16                                                                                       |
| 中     | 中期   | 日本語ドキュメントの補強     | README.ja は良いので、次は troubleshooting、session、run mode、JSON / exit code の日本語チートシートを増やす。3                                                          |
