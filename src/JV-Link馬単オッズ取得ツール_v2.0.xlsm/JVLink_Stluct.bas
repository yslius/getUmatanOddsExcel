Attribute VB_Name = "JVLink_Stluct"
Option Base 0


'  JRA-VAN Data Lab. JV-Data構造体
'
'
'   作成: JRA-VAN ソフトウェア工房
'
'========================================================================
'   (C) Copyright Turf Media System Co.,Ltd. 2003 All rights reserved
'========================================================================


    '''''''''''''''''''' 共通構造体 ''''''''''''''''''''

 '<年月日>
 Private Type YMD
     Year   As String                    ''年
     Month  As String                    ''月
     Day    As String                    ''日
 End Type


 '<時分秒>
 Private Type HMS
     Hour   As String                    ''時
     Minute As String                    ''分
     Second As String                    ''秒
 End Type


 '<時分>
 Private Type HM
     Hour As String                      ''時
     Minute As String                    ''分
 End Type


 '<月日時分>
 Private Type MDHM
     Month As String                     ''月
     Day As String                       ''日
     Hour As String                      ''時
     Minute As String                    ''分
 End Type


 '<レコードヘッダ>
 Private Type RECORD_ID
     RecordSpec As String                ''レコード種別
     DataKubun As String                 ''データ区分
     MakeDate As YMD                     ''データ作成年月日
 End Type


 '<競走識別情報１>
 Private Type RACE_ID
     Year As String                      ''開催年
     MonthDay As String                  ''開催月日
     JyoCD As String                     ''競馬場コード
     Kaiji As String                     ''開催回[第N回]
     Nichiji As String                   ''開催日目[N日目]
     racenum As String                   ''レース番号
 End Type


 '<競走識別情報２>
 Private Type RACE_ID2
     Year As String                      ''開催年
     MonthDay As String                  ''開催月日
     JyoCD As String                     ''競馬場コード
     Kaiji As String                     ''開催回[第N回]
     Nichiji As String                   ''開催日目[N日目]
 End Type


 '<着回数（サイズ3byte）>
 Private Type CHAKUKAISU3_INFO
     Chakukaisu(5) As String
 End Type


 '<着回数（サイズ6byte）>
 Private Type CHAKUKAISU6_INFO
     Chakukaisu(5) As String
 End Type


 '<本年・累計成績情報>
 Private Type SEI_RUIKEI_INFO
     SetYear As String                   ''設定年
     HonSyokinTotal As String            ''本賞金合計
     Fukasyokin As String                ''付加賞金合計
     Chakukaisu(5) As String             ''着回数
 End Type


 '<最近重賞勝利情報>
 Private Type SAIKIN_JYUSYO_INFO
     SaikinJyusyoid As RACE_ID           ''<年月日場回日R>
     Hondai As String                    ''競走名本題
     Ryakusyo10 As String                ''競走名略称10字
     Ryakusyo6 As String                 ''競走名略称6字
     Ryakusyo3 As String                 ''競走名略称3字
     GradeCD As String                   ''グレードコード
     SyussoTosu As String                ''出走頭数
     KettoNum As String                  ''血統登録番号
     Bamei As String                     ''馬名
 End Type


 '<本年・前年・累計成績情報>
 Private Type HON_ZEN_RUIKEISEI_INFO
     SetYear As String                          ''設定年
     HonSyokinHeichi As String                  ''平地本賞金合計
     HonSyokinSyogai As String                  ''障害本賞金合計
     FukaSyokinHeichi As String                 ''平地付加賞金合計
     FukaSyokinSyogai As String                 ''障害付加賞金合計
     ChakuKaisuHeichi As CHAKUKAISU6_INFO       ''平地着回数
     ChakuKaisuSyogai As CHAKUKAISU6_INFO       ''障害着回数
     ChakuKaisuJyo(19) As CHAKUKAISU6_INFO      ''競馬場別着回数
     ChakuKaisuKyori(5) As CHAKUKAISU6_INFO     ''距離別着回数
 End Type


 '<レース情報>
 Private Type RACE_INFO
     YoubiCD As String                   ''曜日コード
     TokuNum As String                   ''特別競走番号
     Hondai As String                    ''競走名本題
     Fukudai As String                   ''競走名副題
     Kakko As String                     ''競走名カッコ内
     HondaiEng As String                 ''競走名本題欧字
     FukudaiEng As String                ''競走名副題欧字
     KakkoEng As String                  ''競走名カッコ内欧字
     Ryakusyo10 As String                ''競走名略称１０字
     Ryakusyo6 As String                 ''競走名略称６字
     Ryakusyo3 As String                 ''競走名略称３字
     Kubun As String                     ''競走名区分
     Nkai As String                      ''重賞回次[第N回]
 End Type


 '<天候・馬場状態>
 Private Type TENKO_BABA_INFO
     TenkoCD As String                   ''天候コード
     SibaBabaCD As String                ''芝馬場状態コード
     DirtBabaCD As String                ''ダート馬場状態コード
 End Type


 '<競走条件>
 Private Type RACE_JYOKEN
     SyubetuCD As String                 ''競走種別コード
     KigoCD As String                    ''競走記号コード
     JyuryoCD As String                  ''重量種別コード
     JyokenCD(4) As String               ''競走条件コード
 End Type

 '''''''''''''''''''' データ構造体 ''''''''''''''''''''

'****** １．特別登録馬 ****************************************
 
 '<登録馬毎情報>
 Private Type TOKUUMA_INFO
     Num As String                       ''連番
     KettoNum As String                  ''血統登録番号
     Bamei As String                     ''馬名
     UmaKigoCD As String                 ''馬記号コード
     SexCD As String                     ''性別コード
     TozaiCD As String                   ''調教師東西所属コード
     ChokyosiCode As String              ''調教師コード
     ChokyosiRyakusyo As String          ''調教師名略称
     Futan As String                     ''負担重量
     Koryu As String                     ''交流区分
 End Type

 Public Type JV_TK_TOKUUMA
     head As RECORD_ID                   ''<レコードヘッダー>
     id As RACE_ID                       ''<競走識別情報１>
     RaceInfo As RACE_INFO               ''<レース情報>
     GradeCD As String                   ''グレードコード
     JyokenInfo As RACE_JYOKEN           ''<競走条件コード>
     Kyori As String                     ''距離
     TrackCD As String                   ''トラックコード
     CourseKubunCD As String             ''コース区分
     HandiDate As YMD                    ''ハンデ発表日
     TorokuTosu As String                ''登録頭数
     TokuUmaInfo(299) As TOKUUMA_INFO    ''<登録馬毎情報>
     crlf As String                      ''レコード区切
     
 End Type

 '****** ２．レース詳細 ****************************************

 '<コーナー通過順位>
 Private Type CORNER_INFO
     Corner As String                    ''コーナー
     Syukaisu As String                  ''周回数
     Jyuni As String                    ''各通過順位
    
 End Type

 Public Type JV_RA_RACE
     head As RECORD_ID                   ''<レコードヘッダー>
     id As RACE_ID                       ''<競走識別情報１>
     RaceInfo As RACE_INFO               ''<レース情報>
     GradeCD As String                   ''グレードコード
     GradeCDBefore As String             ''変更前グレードコード
     JyokenInfo As RACE_JYOKEN           ''<競走条件コード>
     JyokenName As String                ''競走条件名称
     Kyori As String                     ''距離
     KyoriBefore As String               ''変更前距離
     TrackCD As String                    ''トラックコード
     TrackCDBefore As String             ''変更前トラックコード
     CourseKubunCD As String             ''コース区分
     CourseKubunCDBefore As String       ''変更前コース区分
     Honsyokin(6) As String              ''本賞金
     HonsyokinBefore(4) As String        ''変更前本賞金
     Fukasyokin(4) As String             ''付加賞金
     FukasyokinBefore(2) As String       ''変更前付加賞金
     HassoTime As String                 ''発走時刻
     HassoTimeBefore As String           ''変更前発走時刻
     TorokuTosu As String                ''登録頭数
     SyussoTosu As String                ''出走頭数
     NyusenTosu As String                ''入線頭数
     TenkoBaba As TENKO_BABA_INFO        ''天候・馬場状態コード
     LapTime(24) As String               ''ラップタイム
     SyogaiMileTime As String            ''障害マイルタイム
     HaronTimeS3 As String               ''前３ハロンタイム
     HaronTimeS4 As String               ''前４ハロンタイム
     HaronTimeL3 As String               ''後３ハロンタイム
     HaronTimeL4 As String               ''後４ハロンタイム
     CornerInfo(3) As CORNER_INFO        ''<コーナー通過順位>
     RecordUpKubun As String             ''レコード更新区分
     crlf As String                      ''レコード区切り
 End Type


 '****** ３．馬毎レース情報 ****************************************

 '<1着馬(相手馬)情報>
 Private Type CHAKUUMA_INFO
     KettoNum As String                  ''血統登録番号
     Bamei As String                     ''馬名
 End Type

 Public Type JV_SE_RACE_UMA
     head As RECORD_ID                   ''<レコードヘッダー>
     id As RACE_ID                       ''<競走識別情報１>
     Wakuban As String                   ''枠番
     Umaban As String                    ''馬番
     KettoNum As String                  ''血統登録番号
     Bamei As String                     ''馬名
     UmaKigoCD As String                 ''馬記号コード
     SexCD As String                     ''性別コード
     HinsyuCD As String                  ''品種コード
     KeiroCD As String                   ''毛色コード
     Barei As String                     ''馬齢
     TozaiCD As String                   ''東西所属コード
     ChokyosiCode As String              ''調教師コード
     ChokyosiRyakusyo As String          ''調教師名略称
     BanusiCode As String                ''馬主コード
     BanusiName As String                ''馬主名
     Fukusyoku As String                 ''服色標示
     reserved1 As String                 ''予備
     Futan As String                     ''負担重量
     FutanBefore As String               ''変更前負担重量
     Blinker As String                   ''ブリンカー使用区分
     reserved2 As String                 ''予備
     KisyuCode As String                 ''騎手コード
     KisyuCodeBefore As String           ''変更前騎手コード
     KisyuRyakusyo As String             ''騎手名略称
     KisyuRyakusyoBefore As String       ''変更前騎手名略称
     MinaraiCD As String                 ''騎手見習コード
     MinaraiCDBefore As String           ''変更前騎手見習コード
     BaTaijyu As String                  ''馬体重
     ZogenFugo As String                 ''増減符号
     ZogenSa As String                   ''増減差
     IJyoCD As String                    ''異常区分コード
     NyusenJyuni As String               ''入線順位
     KakuteiJyuni As String              ''確定着順
     DochakuKubun As String              ''同着区分
     DochakuTosu As String               ''同着頭数
     Time As String                      ''走破タイム
     ChakusaCD As String                 ''着差コード
     ChakusaCDP As String                ''+着差コード
     ChakusaCDPP As String               ''++着差コード
     Jyuni1c As String                   ''1コーナーでの順位
     Jyuni2c As String                   ''2コーナーでの順位
     Jyuni3c As String                   ''3コーナーでの順位
     Jyuni4c As String                   ''4コーナーでの順位
     Odds As String                      ''単勝オッズ
     Ninki As String                     ''単勝人気順
     Honsyokin As String                 ''獲得本賞金
     Fukasyokin As String                ''獲得付加賞金
     reserved3 As String                 ''予備
     reserved4 As String                 ''予備
     HaronTimeL4 As String               ''後４ハロンタイム
     HaronTimeL3 As String               ''後３ハロンタイム
     ChakuUmaInfo(2) As CHAKUUMA_INFO    ''<1着馬(相手馬)情報>
     TimeDiff As String                  ''タイム差
     RecordUpKubun As String             ''レコード更新区分
     DMKubun As String                   ''マイニング区分
     DMTime As String                    ''マイニング予想走破タイム
     DMGosaP As String                   ''予測誤差(信頼度)＋
     DMGosaM As String                   ''予測誤差(信頼度)－
     DMJyuni As String                   ''マイニング予想順位
     KyakusituKubun As String            ''今回レース脚質判定
     crlf As String                      ''レコード区切り
 End Type


 '****** ４．払戻 ****************************************

 '<払戻情報１ 単・複・枠>
 Private Type PAY_INFO1
     Umaban As String                    ''馬番
     Pay As String                       ''払戻金
     Ninki As String                     ''人気順
 End Type

 '<払戻情報２ 馬連・ワイド・予備・馬単>
 Private Type PAY_INFO2
     Kumi As String                      ''組番
     Pay As String                       ''払戻金
     Ninki As String                     ''人気順
 End Type

 '<払戻情報３ ３連複>
 Private Type PAY_INFO3
     Kumi As String                      ''組番
     Pay As String                       ''払戻金
     Ninki As String                     ''人気順
 End Type

 '<払戻情報４ 予備>
 Private Type PAY_INFO4
     Kumi As String                      ''組番
     Pay As String                       ''払戻金
     Ninki As String                     ''人気順
 End Type

 Public Type JV_HR_PAY
     head As RECORD_ID                   ''<レコードヘッダー>
     id As RACE_ID                       ''<競走識別情報１>
     TorokuTosu As String                ''登録頭数
     SyussoTosu As String                ''出走頭数
     FuseirituFlag(8) As String          ''不成立フラグ
     TokubaraiFlag(8) As String          ''特払フラグ
     HenkanFlag(8) As String             ''返還フラグ
     HenkanUma(27) As String             ''返還馬番情報(馬番01～28)
     HenkanWaku(7) As String             ''返還枠番情報(枠番1～8)
     HenkanDoWaku(7) As String           ''返還同枠情報(枠番1～8)
     PayTansyo(2) As PAY_INFO1           ''<単勝払戻>
     PayFukusyo(4) As PAY_INFO1          ''<複勝払戻>
     PayWakuren(2) As PAY_INFO1          ''<枠連払戻>
     PayUmaren(2) As PAY_INFO2           ''<馬連払戻>
     PayWide(6) As PAY_INFO2             ''<ワイド払戻>
     PayReserved1(2) As PAY_INFO2        ''<予備>
     PayUmatan(5) As PAY_INFO2           ''<馬単払戻>
     PaySanrenpuku(2) As PAY_INFO3       ''<3連複払戻>
     PaySanrentan(5) As PAY_INFO3        ''<3連単払戻>
     crlf As String                      ''レコード区切り
 End Type


 '****** ５．票数（全掛式）****************************************

 '<票数情報１ 単・複・枠>
 Private Type HYO_INFO1
     Umaban As String                    ''馬番
     Hyo As String                       ''票数
     Ninki As String                     ''人気
 End Type

 '<票数情報２ 馬連・ワイド・馬単>
 Private Type HYO_INFO2
     Kumi As String                      ''組番
     Hyo As String                       ''票数
     Ninki As String                     ''人気
 End Type

 '<票数情報３ ３連複票数>
 Private Type HYO_INFO3
     Kumi As String                      ''組番
     Hyo As String                       ''票数
     Ninki As String                     ''人気
 End Type

 '<票数情報４ 予備>
 Private Type HYO_INFO4
     Kumi As String                      ''組番
     Hyo As String                       ''票数
     Ninki As String                     ''人気
 End Type

 Public Type JV_H1_HYOSU_ZENKAKE
     head As RECORD_ID                   ''<レコードヘッダー>
     id As RACE_ID                       ''<競走識別情報１>
     TorokuTosu As String                ''登録頭数
     SyussoTosu As String                ''出走頭数
     HatubaiFlag(6) As String            ''発売フラグ
     FukuChakuBaraiKey As String         ''複勝着払キー
     HenkanUma(27) As String             ''返還馬番情報(馬番01～28)
     HenkanWaku(7) As String             ''返還枠番情報(枠番1～8)
     HenkanDoWaku(7) As String           ''返還同枠情報(枠番1～8)
     HyoTansyo(27) As HYO_INFO1          ''<単勝票数>
     HyoFukusyo(27) As HYO_INFO1         ''<複勝票数>
     HyoWakuren(35) As HYO_INFO1         ''<枠連票数>
     HyoUmaren(152) As HYO_INFO2         ''<馬連票数>
     HyoWide(152) As HYO_INFO2           ''<ワイド票数>
     HyoUmatan(305) As HYO_INFO2         ''<馬単票数>
     HyoSanrenpuku(815) As HYO_INFO3     ''<3連複票数>
     HyoTotal(13) As String              ''票数合計
     crlf As String                      ''レコード区切り
 End Type


 Public Type JV_H6_HYOSU_SANRENTAN
     head As RECORD_ID                   ''<レコードヘッダー>
     id As RACE_ID                       ''<競走識別情報１>
     TorokuTosu As String                ''登録頭数
     SyussoTosu As String                ''出走頭数
     HatubaiFlag As String               ''発売フラグ
     HenkanUma(17) As String             ''返還馬番情報(馬番01～18)
     HyoSanrentan(4895) As HYO_INFO3     ''<3連単票数>
     HyoTotal(2) As String               ''票数合計
     crlf As String                      ''レコード区切り
 End Type

 '****** ６．オッズ（単複枠）****************************************

 '<単勝オッズ>
 Private Type ODDS_TANSYO_INFO
     Umaban As String                    ''馬番
     Odds As String                      ''オッズ
     Ninki As String                     ''人気順
 End Type

 '<複勝オッズ>
 Private Type ODDS_FUKUSYO_INFO
     Umaban As String                    ''馬番
     OddsLow As String                   ''最低オッズ
     OddsHigh As String                  ''最高オッズ
     Ninki As String                     ''人気順
 End Type

 '<枠連オッズ>
 Private Type ODDS_WAKUREN_INFO
     Kumi As String                      ''組
     Odds As String                      ''オッズ
     Ninki As String                     ''人気順
 End Type

 Public Type JV_O1_ODDS_TANFUKUWAKU
     head As RECORD_ID                   ''<レコードヘッダー>
     id As RACE_ID                       ''<競走識別情報１>
     HappyoTime As MDHM                  ''発表月日時分
     TorokuTosu As String                ''登録頭数
     SyussoTosu As String                ''出走頭数
     TansyoFlag As String                ''発売フラグ 単勝
     FukusyoFlag As String               ''発売フラグ 複勝
     WakurenFlag As String               ''発売フラグ　枠連
     FukuChakuBaraiKey As String         ''複勝着払キー
     OddsTansyoInfo(27) As ODDS_TANSYO_INFO    ''<単勝オッズ>
     OddsFukusyoInfo(27) As ODDS_FUKUSYO_INFO  ''<複勝票数オッズ>
     OddsWakurenInfo(35) As ODDS_WAKUREN_INFO  ''<枠連票数オッズ>
     TotalHyosuTansyo As String                ''単勝票数合計
     TotalHyosuFukusyo As String         ''複勝票数合計
     TotalHyosuWakuren As String         ''枠連票数合計
     crlf As String                      ''レコード区切り
 End Type


 '****** ７．オッズ（馬連）****************************************

 '<馬連オッズ>
 Private Type ODDS_UMAREN_INFO
     Kumi As String                      ''組番
     Odds As String                      ''オッズ
     Ninki As String                     ''人気順
 End Type

 Public Type JV_O2_ODDS_UMAREN
     head As RECORD_ID                   ''<レコードヘッダー>
     id As RACE_ID                       ''<競走識別情報１>
     HappyoTime As MDHM                  ''発表月日時分
     TorokuTosu As String                ''登録頭数
     SyussoTosu As String                ''出走頭数
     UmarenFlag As String                ''発売フラグ　馬連
     OddsUmarenInfo(152) As ODDS_UMAREN_INFO   ''<馬連オッズ>
     TotalHyosuUmaren As String          ''馬連票数合計
     crlf As String                      ''レコード区切り
 End Type


 '****** ８．オッズ（ワイド）****************************************

 '<ワイドオッズ>
 Private Type ODDS_WIDE_INFO
     Kumi As String                      ''組番
     OddsLow As String                   ''最低オッズ
     OddsHigh As String                  ''最高オッズ
     Ninki As String                     ''人気順
 End Type

 Public Type JV_O3_ODDS_WIDE
     head As RECORD_ID                   ''<レコードヘッダー>
     id As RACE_ID                       ''<競走識別情報１>
     HappyoTime As MDHM                  ''発表月日時分
     TorokuTosu As String                ''登録頭数
     SyussoTosu As String                ''出走頭数
     WideFlag As String                  ''発売フラグ　ワイド
     OddsWideInfo(152) As ODDS_WIDE_INFO ''<ワイドオッズ>
     TotalHyosuWide As String            ''ワイド票数合計
     crlf As String                      ''レコード区切り
 End Type


 '****** ９．オッズ（馬単） ****************************************

 '<馬単オッズ>
 Private Type ODDS_UMATAN_INFO
     Kumi As String                      ''組番
     Odds As String                      ''オッズ
     Ninki As String                     ''人気順
 End Type

 Public Type JV_O4_ODDS_UMATAN
     head As RECORD_ID                   ''<レコードヘッダー>
     id As RACE_ID                       ''<競走識別情報１>
     HappyoTime As MDHM                  ''発表月日時分
     TorokuTosu As String                ''登録頭数
     SyussoTosu As String                ''出走頭数
     UmatanFlag As String                ''発売フラグ　馬単
     OddsUmatanInfo(305) As ODDS_UMATAN_INFO ''<馬単オッズ>
     TotalHyosuUmatan As String          ''馬単票数合計
     crlf As String                      ''レコード区切り
 End Type


 '****** １０．オッズ（３連複）***************************************

 '<3連複オッズ>
 Private Type ODDS_SANREN_INFO
     Kumi As String                      ''組番
     Odds As String                      ''オッズ
     Ninki As String                     ''人気順
 End Type

 Public Type JV_O5_ODDS_SANREN
     head As RECORD_ID                   ''<レコードヘッダー>
     id As RACE_ID                       ''<競走識別情報１>
     HappyoTime As MDHM                  ''発表月日時分
     TorokuTosu As String                ''登録頭数
     SyussoTosu As String                ''出走頭数
     SanrenpukuFlag As String            ''発売フラグ　3連複
     OddsSanrenInfo(815) As ODDS_SANREN_INFO ''<3連複オッズ>
     TotalHyosuSanrenpuku As String          ''3連複票数合計
     crlf As String                          ''レコード区切り
 End Type


 '****** １０－１．オッズ（３連単）***************************************

 '<3連単オッズ>
 Private Type ODDS_SANRENTAN_INFO
     Kumi As String                      ''組番
     Odds As String                      ''オッズ
     Ninki As String                     ''人気順
 End Type

 Public Type JV_O6_ODDS_SANRENTAN
     head As RECORD_ID                                                          ''<レコードヘッダー>
     id As RACE_ID                                                                      ''<競走識別情報１>
     HappyoTime As MDHM                                                         ''発表月日時分
     TorokuTosu As String                                                       ''登録頭数
     SyussoTosu As String                                                       ''出走頭数
     SanrentanFlag As String                                    ''発売フラグ　3連単
     OddsSanrentanInfo(4895) As ODDS_SANRENTAN_INFO     ''<3連単オッズ>
     TotalHyosuSanrentan As String                                      ''3連単票数合計
     crlf As String                                                                     ''レコード区切り
 End Type
 
  Public Type JV_O6_ODDS_SANRENTAN2
     head As RECORD_ID                                                          ''<レコードヘッダー>
     id As RACE_ID                                                                      ''<競走識別情報１>
     HappyoTime As MDHM                                                         ''発表月日時分
     TorokuTosu As String                                                       ''登録頭数
     SyussoTosu As String                                                       ''出走頭数
     SanrentanFlag As String                                    ''発売フラグ　3連単
     OddsSanrentanInfo As New Collection      ''<3連単オッズ>
     TotalHyosuSanrentan As String                                      ''3連単票数合計
     crlf As String                                                                     ''レコード区切り
 End Type


 '****** １１．競走馬マスタ ****************************************

 '<３代血統情報>
 Private Type KETTO3_INFO
     HansyokuNum As String               ''繁殖登録番号
     Bamei As String                     ''馬名
 End Type

 Public Type JV_UM_UMA
     head As RECORD_ID                   ''<レコードヘッダー>
     KettoNum As String                  ''血統登録番号
     DelKubun As String                  ''競走馬抹消区分
     RegDate As YMD                      ''競走馬登録年月日
     DelDate As YMD                      ''競走馬抹消年月日
     BirthDate As YMD                    ''生年月日
     Bamei As String                     ''馬名
     BameiKana As String                 ''馬名半角カナ
     BameiEng As String                  ''馬名欧字
     UmaKigoCD As String                 ''馬記号コード
     SexCD As String                     ''性別コード
     HinsyuCD As String                  ''品種コード
     KeiroCD As String                   ''毛色コード
     Ketto3Info(13) As KETTO3_INFO       ''<3代血統情報>
     TozaiCD As String                   ''東西所属コード
     ChokyosiCode As String              ''調教師コード
     ChokyosiRyakusyo As String          ''調教師名略称
     Syotai As String                    ''招待地域名
     BreederCode As String               ''生産者コード
     BreederName As String              ''生産者名
     SanchiName As String                ''産地名
     BanusiCode As String                ''馬主コード
     BanusiName As String                ''馬主名
     RuikeiHonsyoHeiti As String         ''平地本賞金累計
     RuikeiHonsyoSyogai As String        ''障害本賞金累計
     RuikeiFukaHeichi As String          ''平地付加賞金累計
     RuikeiFukaSyogai As String          ''障害付加賞金累計
     RuikeiSyutokuHeichi As String       ''平地収得賞金累計
     RuikeiSyutokuSyogai As String       ''障害収得賞金累計
     ChakuSogo As CHAKUKAISU3_INFO       ''総合着回数
     ChakuChuo As CHAKUKAISU3_INFO       ''中央合計着回数
     ChakuKaisuBa(6) As CHAKUKAISU3_INFO ''馬場別着回数
     ChakuKaisuJyotai(11) As CHAKUKAISU3_INFO      ''馬場状態別着回数
     ChakuKaisuKyori(5) As CHAKUKAISU3_INFO        ''距離別着回数
     Kyakusitu(3) As String              ''脚質傾向
     RaceCount As String                 ''登録レース数
     crlf As String                      ''レコード区切り
 End Type


 '****** １２．騎手マスタ ****************************************

 '<初騎乗情報>
 Private Type HATUKIJYO_INFO
     Hatukijyoid As RACE_ID              ''年月日場回日R
     SyussoTosu As String                ''出走頭数
     KettoNum As String                  ''血統登録番号
     Bamei As String                     ''馬名
     KakuteiJyuni As String              ''確定着順
     IJyoCD As String                    ''異常区分コード
 End Type

 '<初勝利情報>
 Private Type HATUSYORI_INFO
     Hatusyoriid As RACE_ID              ''年月日場回日R
     SyussoTosu As String                ''出走頭数
     KettoNum As String                  ''血統登録番号
     Bamei As String                     ''馬名
 End Type

 Public Type JV_KS_KISYU
     head As RECORD_ID                   ''<レコードヘッダー>
     KisyuCode As String                 ''騎手コード
     DelKubun As String                  ''騎手抹消区分
     IssueDate As YMD                    ''騎手免許交付年月日
     DelDate As YMD                      ''騎手免許抹消年月日
     BirthDate As YMD                    ''生年月日
     KisyuName As String                 ''騎手名漢字
     reserved As String                  ''予備
     KisyuNameKana As String             ''騎手名半角カナ
     KisyuRyakusyo As String             ''騎手名略称
     KisyuNameEng As String              ''騎手名欧字
     SexCD As String                     ''性別区分
     SikakuCD As String                  ''騎乗資格コード
     MinaraiCD As String                 ''騎手見習コード
     TozaiCD As String                   ''騎手東西所属コード
     Syotai As String                    ''招待地域名
     ChokyosiCode As String              ''所属調教師コード
     ChokyosiRyakusyo As String          ''所属調教師名略称
     HatuKiJyo(1) As HATUKIJYO_INFO      ''<初騎乗情報>
     HatuSyori(1) As HATUSYORI_INFO      ''<初勝利情報>
     SaikinJyusyo(2) As SAIKIN_JYUSYO_INFO     ''<最近重賞勝利情報>
     HonZenRuikei(2) As HON_ZEN_RUIKEISEI_INFO ''<本年・前年・累計成績情報>
     crlf As String                           ''レコード区切り
 End Type


 '****** １３．調教師マスタ ****************************************

 Public Type JV_CH_CHOKYOSI
     head As RECORD_ID                   ''<レコードヘッダー>
     ChokyosiCode As String              ''調教師コード
     DelKubun As String                  ''調教師抹消区分
     IssueDate As YMD                    ''調教師免許交付年月日
     DelDate As YMD                      ''調教師免許抹消年月日
     BirthDate As YMD                    ''生年月日
     ChokyosiName As String              ''調教師名漢字
     ChokyosiNameKana As String          ''調教師名半角カナ
     ChokyosiRyakusyo As String          ''調教師名略称
     ChokyosiNameEng As String           ''調教師名欧字
     SexCD As String                     ''性別区分
     TozaiCD As String                   ''調教師東西所属コード
     Syotai As String                    ''招待地域名
     SaikinJyusyo(2) As SAIKIN_JYUSYO_INFO     ''<最近重賞勝利情報>
     HonZenRuikei(2) As HON_ZEN_RUIKEISEI_INFO ''<本年・前年・累計成績情報>
     crlf As String                      ''レコード区切り
 End Type


 '******１４．生産者マスタ ****************************************

 Public Type JV_BR_BREEDER
     head As RECORD_ID                   ''<レコードヘッダー>
     BreederCode As String               ''生産者コード
     BreederName_Co As String            ''生産者名（法人格有）
     BreederName As String               ''生産者名（法人格無）
     BreederNameKana As String           ''生産者名半角カナ
     BreederNameEng As String            ''生産者名欧字
     Address As String                   ''生産者住所自治省名
     HonRuikei(1) As SEI_RUIKEI_INFO     ''<本年・累計成績情報>
     crlf As String                      ''レコード区切り
 End Type


 '****** １５．馬主マスタ ****************************************

 Public Type JV_BN_BANUSI
     head As RECORD_ID                   ''<レコードヘッダー>
     BanusiCode As String                ''馬主コード
     BanusiName_Co As String             ''馬主名（法人格有）
     BanusiName As String                ''馬主名（法人格無）
     BanusiNameKana As String            ''馬主名半角カナ
     BanusiNameEng As String             ''馬主名欧字
     Fukusyoku As String                 ''服色標示
     HonRuikei(1) As SEI_RUIKEI_INFO     ''<本年・累計成績情報>
     crlf As String                      ''レコード区切り
 End Type


 '****** １６．繁殖馬マスタ ****************************************

 Public Type JV_HN_HANSYOKU
     head As RECORD_ID                   ''<レコードヘッダー>
     HansyokuNum As String               ''繁殖登録番号
     reserved As String                  ''予備
     KettoNum As String                  ''血統登録番号
     DelKubun As String                  ''繁殖馬抹消区分
     Bamei As String                     ''馬名
     BameiKana As String                 ''馬名半角カナ
     BameiEng As String                  ''馬名欧字
     BirthYear As String                 ''生年
     SexCD As String                     ''性別コード
     HinsyuCD As String                  ''品種コード
     KeiroCD As String                   ''毛色コード
     HansyokuMochiKubun As String        ''繁殖馬持込区分
     ImportYear As String                ''輸入年
     SanchiName As String                ''産地名
     HansyokuFNum As String              ''父馬繁殖登録番号
     HansyokuMNum As String              ''母馬繁殖登録番号
     crlf As String                      ''レコード区切り
 End Type


 '****** １７．産駒マスタ ****************************************

 Public Type JV_SK_SANKU
     head As RECORD_ID                   ''<レコードヘッダー>
     KettoNum As String                  ''血統登録番号
     BirthDate As YMD                    ''生年月日
     SexCD As String                     ''性別コード
     HinsyuCD As String                  ''品種コード
     KeiroCD As String                   ''毛色コード
     SankuMochiKubun As String           ''産駒持込区分
     ImportYear As String                ''輸入年
     BreederCode As String               ''生産者コード
     SanchiName As String                ''産地名
     HansyokuNum(13) As String           ''3代血統 繁殖登録番号
     crlf As String                      ''レコード区切り
 End Type


 '****** １８．レコードマスタ ****************************************

 '<レコード保持馬情報>
 Private Type RECUMA_INFO
     KettoNum As String                  ''血統登録番号
     Bamei As String                     ''馬名
     UmaKigoCD As String                 ''馬記号コード
     SexCD As String                     ''性別コード
     ChokyosiCode As String              ''調教師コード
     ChokyosiName As String              ''調教師名
     Futan As String                     ''負担重量
     KisyuCode As String                 ''騎手コード
     KisyuName As String                 ''騎手名
 End Type

 Public Type JV_RC_RECORD
     head As RECORD_ID                   ''<レコードヘッダー>
     RecInfoKubun As String              ''レコード識別区分
     id As RACE_ID                       ''<競走識別情報１>
     TokuNum As String                   ''特別競走番号
     Hondai As String                    ''競走名本題
     GradeCD As String                   ''グレードコード
     SyubetuCD As String                 ''競走種別コード
     Kyori As String                     ''距離
     TrackCD As String                   ''トラックコード
     RecKubun As String                  ''レコード区分
     RecTime As String                   ''レコードタイム
     TenkoBaba As TENKO_BABA_INFO        ''天候・馬場状態
     RecUmaInfo(2) As RECUMA_INFO        ''<レコード保持馬情報>
     crlf As String                      ''レコード区切り
 End Type


 '****** １９．坂路調教 ****************************************

 Public Type JV_HC_HANRO
     head As RECORD_ID                   ''<レコードヘッダー>
     TresenKubun As String               ''トレセン区分
     ChokyoDate As YMD                   ''調教年月日
     ChokyoTime As String                ''調教時刻
     KettoNum As String                  ''血統登録番号
     HaronTime4 As String                ''4ハロンタイム合計(800M-0M)
     LapTime4 As String                  ''ラップタイム(800M-600M)
     HaronTime3 As String                ''3ハロンタイム合計(600M-0M)
     LapTime3 As String                  ''ラップタイム(600M-400M)
     HaronTime2 As String                ''2ハロンタイム合計(400M-0M)
     LapTime2 As String                  ''ラップタイム(400M-200M)
     LapTime1 As String                  ''ラップタイム(200M-0M)
     crlf As String                      ''レコード区切り
 End Type


 '****** ２０．馬体重 ****************************************

 '<馬体重情報>
 Private Type BATAIJYU_INFO
     Umaban As String                    ''馬番
     Bamei As String                     ''馬名
     BaTaijyu As String                  ''馬体重
     ZogenFugo As String                 ''増減符号
     ZogenSa As String                   ''増減差
 End Type

 Public Type JV_WH_BATAIJYU
     head As RECORD_ID                   ''<レコードヘッダー>
     id As RACE_ID                       ''<競走識別情報１>
     HappyoTime As MDHM                  ''発表月日時分
     BataijyuInfo(17) As BATAIJYU_INFO   ''<馬体重情報>
     crlf As String                      ''レコード区切り
 End Type


 '****** ２１．天候馬場状態 ******************************************

 Public Type JV_WE_WEATHER
     head As RECORD_ID                   ''<レコードヘッダー>
     id As RACE_ID2                      ''<競走識別情報２>
     HappyoTime As MDHM                  ''発表月日時分
     HenkoID As String                   ''変更識別
     TenkoBaba As TENKO_BABA_INFO        ''現在状態情報
     TenkoBabaBefore As TENKO_BABA_INFO  ''変更前状態情報
     crlf As String                      ''レコード区切り
    
 End Type

 '****** ２２．出走取消・競争除外 ****************************************

 Public Type JV_AV_INFO
     head As RECORD_ID                   ''<レコードヘッダー>
     id As RACE_ID                       ''<競走識別情報１>
     HappyoTime As MDHM                  ''発表月日時分
     Umaban As String                    ''馬番
     Bamei As String                     ''馬名
     JiyuKubun As String                 ''事由区分
     crlf As String                      ''レコード区切り
   
 End Type

 '************ ２３．騎手変更 ****************************************

 '<変更情報>
 Private Type JC_INFO
     Futan As String                     ''負担重量
     KisyuCode As String                 ''騎手コード
     KisyuName As String                 ''騎手名
     MinaraiCD As String                 ''騎手見習コード
    
 End Type

 Public Type JV_JC_INFO
     head As RECORD_ID                   ''<レコードヘッダー>
     id As RACE_ID                       ''<競走識別情報１>
     HappyoTime As MDHM                  ''発表月日時分
     Umaban As String                    ''馬番
     Bamei As String                     ''馬名
     JCInfoAfter As JC_INFO              ''<変更後情報>
     JCInfoBefore As JC_INFO             ''<変更前情報>
     crlf As String                      ''レコード区切り
 End Type


 '************ ２３－１．発走時刻変更 ****************************************

 '<変更情報>
 Private Type TC_INFO
     Ji As String                                               ''時
     Fun As String                                              ''分
 End Type

 Public Type JV_TC_INFO
     head As RECORD_ID                   ''<レコードヘッダー>
     id As RACE_ID                       ''<競走識別情報１>
     HappyoTime As MDHM                  ''発表月日時分
     TCInfoAfter As TC_INFO              ''<変更後情報>
     TCInfoBefore As TC_INFO             ''<変更前情報>
     crlf As String                      ''レコード区切り
 End Type


 '************ ２３－２．コース変更 ****************************************

 '<変更情報>
 Private Type CC_INFO
     Kyori As String                                    ''距離
     TruckCD As String                                  ''トラックコード
 End Type

 Public Type JV_CC_INFO
     head As RECORD_ID                   ''<レコードヘッダー>
     id As RACE_ID                       ''<競走識別情報１>
     HappyoTime As MDHM                  ''発表月日時分
     CCInfoAfter As CC_INFO              ''<変更後情報>
     CCInfoBefore As CC_INFO             ''<変更前情報>
     JiyuCD As String                    ''事由コード
     crlf As String                      ''レコード区切り
 End Type


 '****** ２４．データマイニング予想***********************************

 '<マイニング予想>
 Private Type DM_INFO
     Umaban As String                    ''馬番
     DMTime As String                    ''予想走破タイム
     DMGosaP As String                   ''予想誤差(信頼度)＋
     DMGosaM As String                   ''予想誤差(信頼度)－
 End Type

 Public Type JV_DM_INFO
     head As RECORD_ID                   ''<レコードヘッダー>
     id As RACE_ID                       ''<競走識別情報１>
     MakeHM As HM                        ''データ作成時分
     DMInfo(17) As DM_INFO               ''<マイニング予想>
     crlf As String                      ''レコード区切り
 End Type


 '****** ２５．開催スケジュール************************************

 '<重賞案内>
 Private Type JYUSYO_INFO
     TokuNum As String                   ''特別競走番号
     Hondai As String                    ''競走名本題
     Ryakusyo10 As String                ''競走名略称10字
     Ryakusyo6 As String                 ''競走名略称6字
     Ryakusyo3 As String                 ''競走名略称3字
     Nkai As String                      ''重賞回次[第N回]
     GradeCD As String                   ''グレードコード
     SyubetuCD As String                 ''競走種別コード
     KigoCD As String                    ''競走記号コード
     JyuryoCD As String                  ''重量種別コード
     Kyori As String                     ''距離
     TrackCD As String                   ''トラックコード
 End Type

 Public Type JV_YS_SCHEDULE
     head As RECORD_ID                   ''<レコードヘッダー>
     id As RACE_ID2                      ''<競走識別情報２>
     YoubiCD As String                   ''曜日コード
     JyusyoInfo(2) As JYUSYO_INFO        ''<重賞案内>
     crlf As String                      ''レコード区切り
 End Type
 
     '''''''''''''''''''' データセット関数 '''''''''''''''''''''''''''
    
   '****** １．特別登録馬 ****************************************
    
    Public Sub SetData_TK(ByRef lBuf As String, ByRef mBuf As JV_TK_TOKUUMA)
    Dim bytBuf() As Byte                            '' Byte列で処理するためのバッファ
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)              '' レコード種別
            .DataKubun = IncMid(bytBuf, p, 1)               '' データ区分
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)                '' 年
                .Month = IncMid(bytBuf, p, 2)               '' 月
                .Day = IncMid(bytBuf, p, 2)                 '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)                    '' 開催年
            .MonthDay = IncMid(bytBuf, p, 4)                '' 開催月日
            .JyoCD = IncMid(bytBuf, p, 2)                   '' 競馬場コード
            .Kaiji = IncMid(bytBuf, p, 2)                   '' 開催回[第N回]
            .Nichiji = IncMid(bytBuf, p, 2)                 '' 開催日目[N日目]
            .racenum = IncMid(bytBuf, p, 2)                 '' レース番号
        End With ' id
        With .RaceInfo
            .YoubiCD = IncMid(bytBuf, p, 1)                 '' 曜日コード
            .TokuNum = IncMid(bytBuf, p, 4)                 '' 特別競走番号
            .Hondai = IncMid(bytBuf, p, 60)                 '' 競走名本題
            .Fukudai = IncMid(bytBuf, p, 60)                '' 競走名副題
            .Kakko = IncMid(bytBuf, p, 60)                  '' 競走名カッコ内
            .HondaiEng = IncMid(bytBuf, p, 120)             '' 競走名本題欧字
            .FukudaiEng = IncMid(bytBuf, p, 120)            '' 競走名副題欧字
            .KakkoEng = IncMid(bytBuf, p, 120)              '' 競走名カッコ内欧字
            .Ryakusyo10 = IncMid(bytBuf, p, 20)             '' 競走名略称１０字
            .Ryakusyo6 = IncMid(bytBuf, p, 12)              '' 競走名略称６字
            .Ryakusyo3 = IncMid(bytBuf, p, 6)               '' 競走名略称３字
            .Kubun = IncMid(bytBuf, p, 1)                   '' 競走名区分
            .Nkai = IncMid(bytBuf, p, 3)                    '' 重賞回次[第N回]
        End With ' RaceInfo
        .GradeCD = IncMid(bytBuf, p, 1)                     '' グレードコード
        With .JyokenInfo
            .SyubetuCD = IncMid(bytBuf, p, 2)               '' 競走種別コード
            .KigoCD = IncMid(bytBuf, p, 3)                  '' 競走記号コード
            .JyuryoCD = IncMid(bytBuf, p, 1)                '' 重量種別コード
            For j = 0 To 4
                .JyokenCD(j) = IncMid(bytBuf, p, 3)         '' 競走条件コード
            Next j
        End With ' JyokenInfo
        .Kyori = IncMid(bytBuf, p, 4)                       '' 距離
        .TrackCD = IncMid(bytBuf, p, 2)                     '' トラックコード
        .CourseKubunCD = IncMid(bytBuf, p, 2)               '' コース区分
        With .HandiDate
            .Year = IncMid(bytBuf, p, 4)                    '' 年
            .Month = IncMid(bytBuf, p, 2)                   '' 月
            .Day = IncMid(bytBuf, p, 2)                     '' 日
        End With ' HandiDate
        .TorokuTosu = IncMid(bytBuf, p, 3)                  '' 登録頭数
        For i = 0 To 299
            With .TokuUmaInfo(i)
                .Num = IncMid(bytBuf, p, 3)                 '' 連番
                .KettoNum = IncMid(bytBuf, p, 10)           '' 血統登録番号
                .Bamei = IncMid(bytBuf, p, 36)              '' 馬名
                .UmaKigoCD = IncMid(bytBuf, p, 2)           '' 馬記号コード
                .SexCD = IncMid(bytBuf, p, 1)               '' 性別コード
                .TozaiCD = IncMid(bytBuf, p, 1)             '' 調教師東西所属コード
                .ChokyosiCode = IncMid(bytBuf, p, 5)        '' 調教師コード
                .ChokyosiRyakusyo = IncMid(bytBuf, p, 8)    '' 調教師名略称
                .Futan = IncMid(bytBuf, p, 3)               '' 負担重量
                .Koryu = IncMid(bytBuf, p, 1)               '' 交流区分
            End With ' TokuUmaInfo
        Next i
        .crlf = IncMid(bytBuf, p, 2)                        '' レコード区切
    End With

    'バッファ領域解放
    Erase bytBuf
       
    End Sub

    '****** ２．レース詳細 ****************************************
    Public Sub SetData_RA(ByRef lBuf As String, ByRef mBuf As JV_RA_RACE)
    Dim bytBuf() As Byte                            '' Byte列で処理するためのバッファ
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)              '' レコード種別
            .DataKubun = IncMid(bytBuf, p, 1)               '' データ区分
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)                '' 年
                .Month = IncMid(bytBuf, p, 2)               '' 月
                .Day = IncMid(bytBuf, p, 2)                 '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)                    '' 開催年
            .MonthDay = IncMid(bytBuf, p, 4)                '' 開催月日
            .JyoCD = IncMid(bytBuf, p, 2)                   '' 競馬場コード
            .Kaiji = IncMid(bytBuf, p, 2)                   '' 開催回[第N回]
            .Nichiji = IncMid(bytBuf, p, 2)                 '' 開催日目[N日目]
            .racenum = IncMid(bytBuf, p, 2)                 '' レース番号
        End With ' id
        With .RaceInfo
            .YoubiCD = IncMid(bytBuf, p, 1)                 '' 曜日コード
            .TokuNum = IncMid(bytBuf, p, 4)                 '' 特別競走番号
            .Hondai = IncMid(bytBuf, p, 60)                 '' 競走名本題
            .Fukudai = IncMid(bytBuf, p, 60)                '' 競走名副題
            .Kakko = IncMid(bytBuf, p, 60)                  '' 競走名カッコ内
            .HondaiEng = IncMid(bytBuf, p, 120)             '' 競走名本題欧字
            .FukudaiEng = IncMid(bytBuf, p, 120)            '' 競走名副題欧字
            .KakkoEng = IncMid(bytBuf, p, 120)              '' 競走名カッコ内欧字
            .Ryakusyo10 = IncMid(bytBuf, p, 20)             '' 競走名略称１０字
            .Ryakusyo6 = IncMid(bytBuf, p, 12)              '' 競走名略称６字
            .Ryakusyo3 = IncMid(bytBuf, p, 6)               '' 競走名略称３字
            .Kubun = IncMid(bytBuf, p, 1)                   '' 競走名区分
            .Nkai = IncMid(bytBuf, p, 3)                    '' 重賞回次[第N回]
        End With ' RaceInfo
        .GradeCD = IncMid(bytBuf, p, 1)                     '' グレードコード
        .GradeCDBefore = IncMid(bytBuf, p, 1)               '' 変更前グレードコード
        With .JyokenInfo
            .SyubetuCD = IncMid(bytBuf, p, 2)               '' 競走種別コード
            .KigoCD = IncMid(bytBuf, p, 3)                  '' 競走記号コード
            .JyuryoCD = IncMid(bytBuf, p, 1)                '' 重量種別コード
            For j = 0 To 4
                .JyokenCD(j) = IncMid(bytBuf, p, 3)         '' 競走条件コード
            Next j
        End With ' JyokenInfo
        .JyokenName = IncMid(bytBuf, p, 60)                 '' 競走条件名称
        .Kyori = IncMid(bytBuf, p, 4)                       '' 距離
        .KyoriBefore = IncMid(bytBuf, p, 4)                 '' 変更前距離
        .TrackCD = IncMid(bytBuf, p, 2)                     '' トラックコード
        .TrackCDBefore = IncMid(bytBuf, p, 2)               '' 変更前トラックコード
        .CourseKubunCD = IncMid(bytBuf, p, 2)               '' コース区分
        .CourseKubunCDBefore = IncMid(bytBuf, p, 2)         '' 変更前コース区分
        For i = 0 To 6
            .Honsyokin(i) = IncMid(bytBuf, p, 8)            '' 本賞金
        Next i
        For i = 0 To 4
            .HonsyokinBefore(i) = IncMid(bytBuf, p, 8)      '' 変更前本賞金
        Next i
        For i = 0 To 4
            .Fukasyokin(i) = IncMid(bytBuf, p, 8)           '' 付加賞金
        Next i
        For i = 0 To 2
            .FukasyokinBefore(i) = IncMid(bytBuf, p, 8)     '' 変更前付加賞金
        Next i
        .HassoTime = IncMid(bytBuf, p, 4)                   '' 発走時刻
        .HassoTimeBefore = IncMid(bytBuf, p, 4)             '' 変更前発走時刻
        .TorokuTosu = IncMid(bytBuf, p, 2)                  '' 登録頭数
        .SyussoTosu = IncMid(bytBuf, p, 2)                  '' 出走頭数
        .NyusenTosu = IncMid(bytBuf, p, 2)                  '' 入線頭数
        With .TenkoBaba
            .TenkoCD = IncMid(bytBuf, p, 1)                 '' 天候コード
            .SibaBabaCD = IncMid(bytBuf, p, 1)              '' 芝馬場状態コード
            .DirtBabaCD = IncMid(bytBuf, p, 1)              '' ダート馬場状態コード
        End With ' TenkoBaba
        For i = 0 To 24
            .LapTime(i) = IncMid(bytBuf, p, 3)              '' ラップタイム
        Next i
        .SyogaiMileTime = IncMid(bytBuf, p, 4)              '' 障害マイルタイム
        .HaronTimeS3 = IncMid(bytBuf, p, 3)                 '' 前３ハロンタイム
        .HaronTimeS4 = IncMid(bytBuf, p, 3)                 '' 前４ハロンタイム
        .HaronTimeL3 = IncMid(bytBuf, p, 3)                 '' 後３ハロンタイム
        .HaronTimeL4 = IncMid(bytBuf, p, 3)                 '' 後４ハロンタイム
        For i = 0 To 3
            With .CornerInfo(i)
                .Corner = IncMid(bytBuf, p, 1)              '' コーナー
                .Syukaisu = IncMid(bytBuf, p, 1)            '' 周回数
                .Jyuni = IncMid(bytBuf, p, 70)              '' 各通過順位
            End With ' CornerInfo
        Next i
        .RecordUpKubun = IncMid(bytBuf, p, 1)               '' レコード更新区分
        .crlf = IncMid(bytBuf, p, 2)        '' レコード区切り
    End With

    'バッファ領域解放
    Erase bytBuf
     
    End Sub


    '****** ３．馬毎レース情報 ****************************************

    Public Sub SetData_SE(ByRef lBuf As String, ByRef mBuf As JV_SE_RACE_UMA)
    Dim bytBuf() As Byte                            '' Byte列で処理するためのバッファ
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)      '' レコード種別
            .DataKubun = IncMid(bytBuf, p, 1)       '' データ区分
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)        '' 年
                .Month = IncMid(bytBuf, p, 2)       '' 月
                .Day = IncMid(bytBuf, p, 2)         '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)            '' 開催年
            .MonthDay = IncMid(bytBuf, p, 4)        '' 開催月日
            .JyoCD = IncMid(bytBuf, p, 2)           '' 競馬場コード
            .Kaiji = IncMid(bytBuf, p, 2)           '' 開催回[第N回]
            .Nichiji = IncMid(bytBuf, p, 2)         '' 開催日目[N日目]
            .racenum = IncMid(bytBuf, p, 2)         '' レース番号
        End With ' id
        .Wakuban = IncMid(bytBuf, p, 1)             '' 枠番
        .Umaban = IncMid(bytBuf, p, 2)              '' 馬番
        .KettoNum = IncMid(bytBuf, p, 10)           '' 血統登録番号
        .Bamei = IncMid(bytBuf, p, 36)              '' 馬名
        .UmaKigoCD = IncMid(bytBuf, p, 2)           '' 馬記号コード
        .SexCD = IncMid(bytBuf, p, 1)               '' 性別コード
        .HinsyuCD = IncMid(bytBuf, p, 1)            '' 品種コード
        .KeiroCD = IncMid(bytBuf, p, 2)             '' 毛色コード
        .Barei = IncMid(bytBuf, p, 2)               '' 馬齢
        .TozaiCD = IncMid(bytBuf, p, 1)             '' 東西所属コード
        .ChokyosiCode = IncMid(bytBuf, p, 5)        '' 調教師コード
        .ChokyosiRyakusyo = IncMid(bytBuf, p, 8)    '' 調教師名略称
        .BanusiCode = IncMid(bytBuf, p, 6)          '' 馬主コード
        .BanusiName = IncMid(bytBuf, p, 64)         '' 馬主名
        .Fukusyoku = IncMid(bytBuf, p, 60)          '' 服色標示
        .reserved1 = IncMid(bytBuf, p, 60)          '' 予備
        .Futan = IncMid(bytBuf, p, 3)               '' 負担重量
        .FutanBefore = IncMid(bytBuf, p, 3)         '' 変更前負担重量
        .Blinker = IncMid(bytBuf, p, 1)             '' ブリンカー使用区分
        .reserved2 = IncMid(bytBuf, p, 1)           '' 予備
        .KisyuCode = IncMid(bytBuf, p, 5)           '' 騎手コード
        .KisyuCodeBefore = IncMid(bytBuf, p, 5)     '' 変更前騎手コード
        .KisyuRyakusyo = IncMid(bytBuf, p, 8)       '' 騎手名略称
        .KisyuRyakusyoBefore = IncMid(bytBuf, p, 8) '' 変更前騎手名略称
        .MinaraiCD = IncMid(bytBuf, p, 1)           '' 騎手見習コード
        .MinaraiCDBefore = IncMid(bytBuf, p, 1)     '' 変更前騎手見習コード
        .BaTaijyu = IncMid(bytBuf, p, 3)            '' 馬体重
        .ZogenFugo = IncMid(bytBuf, p, 1)           '' 増減符号
        .ZogenSa = IncMid(bytBuf, p, 3)             '' 増減差
        .IJyoCD = IncMid(bytBuf, p, 1)              '' 異常区分コード
        .NyusenJyuni = IncMid(bytBuf, p, 2)         '' 入線順位
        .KakuteiJyuni = IncMid(bytBuf, p, 2)        '' 確定着順
        .DochakuKubun = IncMid(bytBuf, p, 1)        '' 同着区分
        .DochakuTosu = IncMid(bytBuf, p, 1)         '' 同着頭数
        .Time = IncMid(bytBuf, p, 4)                '' 走破タイム
        .ChakusaCD = IncMid(bytBuf, p, 3)           '' 着差コード
        .ChakusaCDP = IncMid(bytBuf, p, 3)          '' +着差コード
        .ChakusaCDPP = IncMid(bytBuf, p, 3)         '' ++着差コード
        .Jyuni1c = IncMid(bytBuf, p, 2)             '' 1コーナーでの順位
        .Jyuni2c = IncMid(bytBuf, p, 2)             '' 2コーナーでの順位
        .Jyuni3c = IncMid(bytBuf, p, 2)             '' 3コーナーでの順位
        .Jyuni4c = IncMid(bytBuf, p, 2)             '' 4コーナーでの順位
        .Odds = IncMid(bytBuf, p, 4)                '' 単勝オッズ
        .Ninki = IncMid(bytBuf, p, 2)               '' 単勝人気順
        .Honsyokin = IncMid(bytBuf, p, 8)           '' 獲得本賞金
        .Fukasyokin = IncMid(bytBuf, p, 8)          '' 獲得付加賞金
        .reserved3 = IncMid(bytBuf, p, 3)           '' 予備
        .reserved4 = IncMid(bytBuf, p, 3)           '' 予備
        .HaronTimeL4 = IncMid(bytBuf, p, 3)         '' 後４ハロンタイム
        .HaronTimeL3 = IncMid(bytBuf, p, 3)         '' 後３ハロンタイム
        For i = 0 To 2
            With .ChakuUmaInfo(i)
                .KettoNum = IncMid(bytBuf, p, 10)   '' 血統登録番号
                .Bamei = IncMid(bytBuf, p, 36)      '' 馬名
            End With ' ChakuUmaInfo
        Next i
        .TimeDiff = IncMid(bytBuf, p, 4)            '' タイム差
        .RecordUpKubun = IncMid(bytBuf, p, 1)       '' レコード更新区分
        .DMKubun = IncMid(bytBuf, p, 1)             '' マイニング区分
        .DMTime = IncMid(bytBuf, p, 5)              '' マイニング予想走破タイム
        .DMGosaP = IncMid(bytBuf, p, 4)             '' 予測誤差(信頼度)＋
        .DMGosaM = IncMid(bytBuf, p, 4)             '' 予測誤差(信頼度)－
        .DMJyuni = IncMid(bytBuf, p, 2)             '' マイニング予想順位
        .KyakusituKubun = IncMid(bytBuf, p, 1)      '' 今回レース脚質判定
        .crlf = IncMid(bytBuf, p, 2)                '' レコード区切り
    End With

    'バッファ領域解放
    Erase bytBuf
   
    End Sub


    '****** ４．払戻 ****************************************

    Public Sub SetData_HR(lBuf As String, ByRef mBuf As JV_HR_PAY)
    Dim bytBuf() As Byte                                    '' バイト配列で処理するためのバッファ
    Dim i As Integer                                        '' ループカウンタ
    Dim j As Integer                                        '' ループカウンタ
    Dim k As Integer                                        '' ループカウンタ
    Dim p As Long                                           '' 切り分け開始位置
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)              '' レコード種別
            .DataKubun = IncMid(bytBuf, p, 1)               '' データ区分
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)                '' 年
                .Month = IncMid(bytBuf, p, 2)               '' 月
                .Day = IncMid(bytBuf, p, 2)                 '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)                    '' 開催年
            .MonthDay = IncMid(bytBuf, p, 4)                '' 開催月日
            .JyoCD = IncMid(bytBuf, p, 2)                   '' 競馬場コード
            .Kaiji = IncMid(bytBuf, p, 2)                   '' 開催回[第N回]
            .Nichiji = IncMid(bytBuf, p, 2)                 '' 開催日目[N日目]
            .racenum = IncMid(bytBuf, p, 2)                 '' レース番号
        End With ' id
        .TorokuTosu = IncMid(bytBuf, p, 2)                  '' 登録頭数
        .SyussoTosu = IncMid(bytBuf, p, 2)                  '' 出走頭数
        For i = 0 To 8
            .FuseirituFlag(i) = IncMid(bytBuf, p, 1)        '' 不成立フラグ
        Next i
        For i = 0 To 8
            .TokubaraiFlag(i) = IncMid(bytBuf, p, 1)        '' 特払フラグ
        Next i
        For i = 0 To 8
            .HenkanFlag(i) = IncMid(bytBuf, p, 1)           '' 返還フラグ
        Next i
        For i = 0 To 27
            .HenkanUma(i) = IncMid(bytBuf, p, 1)            '' 返還馬番情報(馬番01～28)
        Next i
        For i = 0 To 7
            .HenkanWaku(i) = IncMid(bytBuf, p, 1)           '' 返還枠番情報(枠番1～8)
        Next i
        For i = 0 To 7
            .HenkanDoWaku(i) = IncMid(bytBuf, p, 1)         '' 返還同枠情報(枠番1～8)
        Next i
        For i = 0 To 2
            With .PayTansyo(i)
                .Umaban = IncMid(bytBuf, p, 2)              '' 馬番
                .Pay = IncMid(bytBuf, p, 9)                 '' 払戻金
                .Ninki = IncMid(bytBuf, p, 2)               '' 人気順
            End With ' PayTansyo
        Next i
        For i = 0 To 4
            With .PayFukusyo(i)
                .Umaban = IncMid(bytBuf, p, 2)              '' 馬番
                .Pay = IncMid(bytBuf, p, 9)                 '' 払戻金
                .Ninki = IncMid(bytBuf, p, 2)               '' 人気順
            End With ' PayFukusyo
        Next i
        For i = 0 To 2
            With .PayWakuren(i)
                .Umaban = IncMid(bytBuf, p, 2)              '' 馬番
                .Pay = IncMid(bytBuf, p, 9)                 '' 払戻金
                .Ninki = IncMid(bytBuf, p, 2)               '' 人気順
            End With ' PayWakuren
        Next i
        For i = 0 To 2
            With .PayUmaren(i)
                .Kumi = IncMid(bytBuf, p, 4)                '' 組番
                .Pay = IncMid(bytBuf, p, 9)                 '' 払戻金
                .Ninki = IncMid(bytBuf, p, 3)               '' 人気順
            End With ' PayUmaren
        Next i
        For i = 0 To 6
            With .PayWide(i)
                .Kumi = IncMid(bytBuf, p, 4)                '' 組番
                .Pay = IncMid(bytBuf, p, 9)                 '' 払戻金
                .Ninki = IncMid(bytBuf, p, 3)               '' 人気順
            End With ' PayWide
        Next i
        For i = 0 To 2
            With .PayReserved1(i)
                .Kumi = IncMid(bytBuf, p, 4)                '' 組番
                .Pay = IncMid(bytBuf, p, 9)                 '' 払戻金
                .Ninki = IncMid(bytBuf, p, 3)               '' 人気順
            End With ' PayReserved1
        Next i
        For i = 0 To 5
            With .PayUmatan(i)
                .Kumi = IncMid(bytBuf, p, 4)                '' 組番
                .Pay = IncMid(bytBuf, p, 9)                 '' 払戻金
                .Ninki = IncMid(bytBuf, p, 3)               '' 人気順
            End With ' PayUmatan
        Next i
        For i = 0 To 2
            With .PaySanrenpuku(i)
                .Kumi = IncMid(bytBuf, p, 6)                '' 組番
                .Pay = IncMid(bytBuf, p, 9)                 '' 払戻金
                .Ninki = IncMid(bytBuf, p, 3)               '' 人気順
            End With ' PaySanrenpuku
        Next i
        For i = 0 To 5
            With .PaySanrentan(i)
                .Kumi = IncMid(bytBuf, p, 6)                '' 組番
                .Pay = IncMid(bytBuf, p, 9)                 '' 払戻金
                .Ninki = IncMid(bytBuf, p, 4)               '' 人気順
            End With ' PayReserved2
        Next i
        .crlf = IncMid(bytBuf, p, 2)                        '' レコード区切り
    End With
    
    'バッファ領域解放
    Erase bytBuf
    
    End Sub

    '****** ５．票数（全掛式）****************************************

    Public Sub SetData_H1(lBuf As String, ByRef mBuf As JV_H1_HYOSU_ZENKAKE)
    Dim bytBuf() As Byte                            '' Byte列で処理するためのバッファ
    
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)      '' レコード種別
            .DataKubun = IncMid(bytBuf, p, 1)       '' データ区分
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)        '' 年
                .Month = IncMid(bytBuf, p, 2)       '' 月
                .Day = IncMid(bytBuf, p, 2)         '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)            '' 開催年
            .MonthDay = IncMid(bytBuf, p, 4)        '' 開催月日
            .JyoCD = IncMid(bytBuf, p, 2)           '' 競馬場コード
            .Kaiji = IncMid(bytBuf, p, 2)           '' 開催回[第N回]
            .Nichiji = IncMid(bytBuf, p, 2)         '' 開催日目[N日目]
            .racenum = IncMid(bytBuf, p, 2)         '' レース番号
        End With ' id
        .TorokuTosu = IncMid(bytBuf, p, 2)          '' 登録頭数
        .SyussoTosu = IncMid(bytBuf, p, 2)          '' 出走頭数
        For i = 0 To 6
            .HatubaiFlag(i) = IncMid(bytBuf, p, 1)  '' 発売フラグ
        Next i
        .FukuChakuBaraiKey = IncMid(bytBuf, p, 1)   '' 複勝着払キー
        For i = 0 To 27
            .HenkanUma(i) = IncMid(bytBuf, p, 1)    '' 返還馬番情報(馬番01～28)
        Next i
        For i = 0 To 7
            .HenkanWaku(i) = IncMid(bytBuf, p, 1)   '' 返還枠番情報(枠番1～8)
        Next i
        For i = 0 To 7
            .HenkanDoWaku(i) = IncMid(bytBuf, p, 1) '' 返還同枠情報(枠番1～8)
        Next i
        For i = 0 To 27
            With .HyoTansyo(i)
                .Umaban = IncMid(bytBuf, p, 2)      '' 馬番
                .Hyo = IncMid(bytBuf, p, 11)        '' 票数
                .Ninki = IncMid(bytBuf, p, 2)       '' 人気
            End With ' HyoTansyo
        Next i
        For i = 0 To 27
            With .HyoFukusyo(i)
                .Umaban = IncMid(bytBuf, p, 2)      '' 馬番
                .Hyo = IncMid(bytBuf, p, 11)        '' 票数
                .Ninki = IncMid(bytBuf, p, 2)       '' 人気
            End With ' HyoFukusyo
        Next i
        For i = 0 To 35
            With .HyoWakuren(i)
                .Umaban = IncMid(bytBuf, p, 2)      '' 馬番
                .Hyo = IncMid(bytBuf, p, 11)        '' 票数
                .Ninki = IncMid(bytBuf, p, 2)       '' 人気
            End With ' HyoWakuren
        Next i
        For i = 0 To 152
            With .HyoUmaren(i)
                .Kumi = IncMid(bytBuf, p, 4)        '' 組番
                .Hyo = IncMid(bytBuf, p, 11)        '' 票数
                .Ninki = IncMid(bytBuf, p, 3)       '' 人気
            End With ' HyoUmaren
        Next i
        For i = 0 To 152
            With .HyoWide(i)
                .Kumi = IncMid(bytBuf, p, 4)        '' 組番
                .Hyo = IncMid(bytBuf, p, 11)        '' 票数
                .Ninki = IncMid(bytBuf, p, 3)       '' 人気
            End With ' HyoWide
        Next i
        For i = 0 To 305
            With .HyoUmatan(i)
                .Kumi = IncMid(bytBuf, p, 4)        '' 組番
                .Hyo = IncMid(bytBuf, p, 11)        '' 票数
                .Ninki = IncMid(bytBuf, p, 3)       '' 人気
            End With ' HyoUmatan
        Next i
        For i = 0 To 815
            With .HyoSanrenpuku(i)
                .Kumi = IncMid(bytBuf, p, 6)        '' 組番
                .Hyo = IncMid(bytBuf, p, 11)        '' 票数
                .Ninki = IncMid(bytBuf, p, 3)       '' 人気
            End With ' HyoSanrenpuku
        Next i
        For i = 0 To 13
            .HyoTotal(i) = IncMid(bytBuf, p, 11)    '' 票数合計
        Next i
        .crlf = IncMid(bytBuf, p, 2)                '' レコード区切り
    End With

    'バッファ領域解放
    Erase bytBuf
    
    End Sub


    '****** ６．オッズ（単複枠）****************************************

    Public Sub SetData_O1(lBuf As String, ByRef mBuf As JV_O1_ODDS_TANFUKUWAKU)
    Dim bytBuf() As Byte                            '' Byte列で処理するためのバッファ
    
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)      '' レコード種別
            .DataKubun = IncMid(bytBuf, p, 1)       '' データ区分
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)        '' 年
                .Month = IncMid(bytBuf, p, 2)       '' 月
                .Day = IncMid(bytBuf, p, 2) '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)            '' 開催年
            .MonthDay = IncMid(bytBuf, p, 4)        '' 開催月日
            .JyoCD = IncMid(bytBuf, p, 2)           '' 競馬場コード
            .Kaiji = IncMid(bytBuf, p, 2)           '' 開催回[第N回]
            .Nichiji = IncMid(bytBuf, p, 2)         '' 開催日目[N日目]
            .racenum = IncMid(bytBuf, p, 2)         '' レース番号
        End With ' id
        With .HappyoTime
            .Month = IncMid(bytBuf, p, 2)           '' 月
            .Day = IncMid(bytBuf, p, 2)             '' 日
            .Hour = IncMid(bytBuf, p, 2)            '' 時
            .Minute = IncMid(bytBuf, p, 2)          '' 分
        End With ' HappyoTime
        .TorokuTosu = IncMid(bytBuf, p, 2)          '' 登録頭数
        .SyussoTosu = IncMid(bytBuf, p, 2)          '' 出走頭数
        .TansyoFlag = IncMid(bytBuf, p, 1)          '' 発売フラグ
        .FukusyoFlag = IncMid(bytBuf, p, 1)         '' 発売フラグ
        .WakurenFlag = IncMid(bytBuf, p, 1)         '' 発売フラグ　枠連
        .FukuChakuBaraiKey = IncMid(bytBuf, p, 1)   '' 複勝着払キー
        For i = 0 To 27
            With .OddsTansyoInfo(i)
                .Umaban = IncMid(bytBuf, p, 2)      '' 馬番
                .Odds = IncMid(bytBuf, p, 4)        '' オッズ
                .Ninki = IncMid(bytBuf, p, 2)       '' 人気順
            End With ' OddsTansyoInfo
        Next i
        For i = 0 To 27
            With .OddsFukusyoInfo(i)
                .Umaban = IncMid(bytBuf, p, 2)      '' 馬番
                .OddsLow = IncMid(bytBuf, p, 4)     '' 最低オッズ
                .OddsHigh = IncMid(bytBuf, p, 4)    '' 最高オッズ
                .Ninki = IncMid(bytBuf, p, 2)       '' 人気順
            End With ' OddsFukusyoInfo
        Next i
        For i = 0 To 35
            With .OddsWakurenInfo(i)
                .Kumi = IncMid(bytBuf, p, 2)        '' 組
                .Odds = IncMid(bytBuf, p, 5)        '' オッズ
                .Ninki = IncMid(bytBuf, p, 2)       '' 人気順
            End With ' OddsWakurenInfo
        Next i
        .TotalHyosuTansyo = IncMid(bytBuf, p, 11)   '' 単勝票数合計
        .TotalHyosuFukusyo = IncMid(bytBuf, p, 11)  '' 複勝票数合計
        .TotalHyosuWakuren = IncMid(bytBuf, p, 11)  '' 枠連票数合計
        .crlf = IncMid(bytBuf, p, 2)                '' レコード区切り
    End With

    'バッファ領域解放
    Erase bytBuf

    End Sub


    '****** ７．オッズ（馬連）****************************************

    Public Sub SetData_O2(lBuf As String, ByRef mBuf As JV_O2_ODDS_UMAREN)
    Dim bytBuf() As Byte                            '' Byte列で処理するためのバッファ
    
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)      '' レコード種別
            .DataKubun = IncMid(bytBuf, p, 1)       '' データ区分
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)        '' 年
                .Month = IncMid(bytBuf, p, 2)       '' 月
                .Day = IncMid(bytBuf, p, 2) '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)    '' 開催年
            .MonthDay = IncMid(bytBuf, p, 4)        '' 開催月日
            .JyoCD = IncMid(bytBuf, p, 2)   '' 競馬場コード
            .Kaiji = IncMid(bytBuf, p, 2)   '' 開催回[第N回]
            .Nichiji = IncMid(bytBuf, p, 2) '' 開催日目[N日目]
            .racenum = IncMid(bytBuf, p, 2) '' レース番号
        End With ' id
        With .HappyoTime
            .Month = IncMid(bytBuf, p, 2)   '' 月
            .Day = IncMid(bytBuf, p, 2)     '' 日
            .Hour = IncMid(bytBuf, p, 2)    '' 時
            .Minute = IncMid(bytBuf, p, 2)  '' 分
        End With ' HappyoTime
        .TorokuTosu = IncMid(bytBuf, p, 2)  '' 登録頭数
        .SyussoTosu = IncMid(bytBuf, p, 2)  '' 出走頭数
        .UmarenFlag = IncMid(bytBuf, p, 1)  '' 発売フラグ　馬連
        For i = 0 To 152
            With .OddsUmarenInfo(i)
                .Kumi = IncMid(bytBuf, p, 4)        '' 組番
                .Odds = IncMid(bytBuf, p, 6)        '' オッズ
                .Ninki = IncMid(bytBuf, p, 3)       '' 人気順
            End With ' OddsUmarenInfo
        Next i
        .TotalHyosuUmaren = IncMid(bytBuf, p, 11)   '' 馬連票数合計
        .crlf = IncMid(bytBuf, p, 2)        '' レコード区切り
    End With

    'バッファ領域解放
    Erase bytBuf

    End Sub


    '****** ８．オッズ（ワイド）****************************************

    Public Sub SetData_O3(lBuf As String, ByRef mBuf As JV_O3_ODDS_WIDE)
    Dim bytBuf() As Byte                            '' Byte列で処理するためのバッファ
    
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)      '' レコード種別
            .DataKubun = IncMid(bytBuf, p, 1)       '' データ区分
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)        '' 年
                .Month = IncMid(bytBuf, p, 2)       '' 月
                .Day = IncMid(bytBuf, p, 2) '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)    '' 開催年
            .MonthDay = IncMid(bytBuf, p, 4)        '' 開催月日
            .JyoCD = IncMid(bytBuf, p, 2)   '' 競馬場コード
            .Kaiji = IncMid(bytBuf, p, 2)   '' 開催回[第N回]
            .Nichiji = IncMid(bytBuf, p, 2) '' 開催日目[N日目]
            .racenum = IncMid(bytBuf, p, 2) '' レース番号
        End With ' id
        With .HappyoTime
            .Month = IncMid(bytBuf, p, 2)   '' 月
            .Day = IncMid(bytBuf, p, 2)     '' 日
            .Hour = IncMid(bytBuf, p, 2)    '' 時
            .Minute = IncMid(bytBuf, p, 2)  '' 分
        End With ' HappyoTime
        .TorokuTosu = IncMid(bytBuf, p, 2)  '' 登録頭数
        .SyussoTosu = IncMid(bytBuf, p, 2)  '' 出走頭数
        .WideFlag = IncMid(bytBuf, p, 1)    '' 発売フラグ　ワイド
        For i = 0 To 152
            With .OddsWideInfo(i)
                .Kumi = IncMid(bytBuf, p, 4)        '' 組番
                .OddsLow = IncMid(bytBuf, p, 5)     '' 最低オッズ
                .OddsHigh = IncMid(bytBuf, p, 5)    '' 最高オッズ
                .Ninki = IncMid(bytBuf, p, 3)       '' 人気順
            End With ' OddsWideInfo
        Next i
        .TotalHyosuWide = IncMid(bytBuf, p, 11)     '' ワイド票数合計
        .crlf = IncMid(bytBuf, p, 2)        '' レコード区切り
    End With

    'バッファ領域解放
    Erase bytBuf
   
    End Sub


    '****** ９．オッズ（馬単） ****************************************

    Public Sub SetData_O4(lBuf As String, ByRef mBuf As JV_O4_ODDS_UMATAN)
    Dim bytBuf() As Byte                            '' Byte列で処理するためのバッファ
    
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)      '' レコード種別
            .DataKubun = IncMid(bytBuf, p, 1)       '' データ区分
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)        '' 年
                .Month = IncMid(bytBuf, p, 2)       '' 月
                .Day = IncMid(bytBuf, p, 2)         '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)            '' 開催年
            .MonthDay = IncMid(bytBuf, p, 4)        '' 開催月日
            .JyoCD = IncMid(bytBuf, p, 2)           '' 競馬場コード
            .Kaiji = IncMid(bytBuf, p, 2)           '' 開催回[第N回]
            .Nichiji = IncMid(bytBuf, p, 2)         '' 開催日目[N日目]
            .racenum = IncMid(bytBuf, p, 2)         '' レース番号
        End With ' id
        With .HappyoTime
            .Month = IncMid(bytBuf, p, 2)           '' 月
            .Day = IncMid(bytBuf, p, 2)             '' 日
            .Hour = IncMid(bytBuf, p, 2)            '' 時
            .Minute = IncMid(bytBuf, p, 2)          '' 分
        End With ' HappyoTime
        .TorokuTosu = IncMid(bytBuf, p, 2)          '' 登録頭数
        .SyussoTosu = IncMid(bytBuf, p, 2)          '' 出走頭数
        .UmatanFlag = IncMid(bytBuf, p, 1)          '' 発売フラグ　馬単
        For i = 0 To 305
            With .OddsUmatanInfo(i)
                .Kumi = IncMid(bytBuf, p, 4)        '' 組番
                .Odds = IncMid(bytBuf, p, 6)        '' オッズ
                .Ninki = IncMid(bytBuf, p, 3)       '' 人気順
            End With ' OddsUmatanInfo
        Next i
        .TotalHyosuUmatan = IncMid(bytBuf, p, 11)   '' 馬単票数合計
        .crlf = IncMid(bytBuf, p, 2)                '' レコード区切り
    End With

    'バッファ領域解放
    Erase bytBuf

    End Sub


    '****** １０．オッズ（３連複）***************************************

    Public Sub SetData_O5(lBuf As String, ByRef mBuf As JV_O5_ODDS_SANREN)
    Dim bytBuf() As Byte                            '' Byte列で処理するためのバッファ
    
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)      '' レコード種別
            .DataKubun = IncMid(bytBuf, p, 1)       '' データ区分
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)        '' 年
                .Month = IncMid(bytBuf, p, 2)       '' 月
                .Day = IncMid(bytBuf, p, 2)         '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)            '' 開催年
            .MonthDay = IncMid(bytBuf, p, 4)        '' 開催月日
            .JyoCD = IncMid(bytBuf, p, 2)           '' 競馬場コード
            .Kaiji = IncMid(bytBuf, p, 2)               '' 開催回[第N回]
            .Nichiji = IncMid(bytBuf, p, 2)         '' 開催日目[N日目]
            .racenum = IncMid(bytBuf, p, 2)         '' レース番号
        End With ' id
        With .HappyoTime
            .Month = IncMid(bytBuf, p, 2)           '' 月
            .Day = IncMid(bytBuf, p, 2)             '' 日
            .Hour = IncMid(bytBuf, p, 2)            '' 時
            .Minute = IncMid(bytBuf, p, 2)          '' 分
        End With ' HappyoTime
        .TorokuTosu = IncMid(bytBuf, p, 2)          '' 登録頭数
        .SyussoTosu = IncMid(bytBuf, p, 2)          '' 出走頭数
        .SanrenpukuFlag = IncMid(bytBuf, p, 1)      '' 発売フラグ　3連複
        For i = 0 To 815
            With .OddsSanrenInfo(i)
                .Kumi = IncMid(bytBuf, p, 6)        '' 組番
                .Odds = IncMid(bytBuf, p, 6)        '' オッズ
                .Ninki = IncMid(bytBuf, p, 3)       '' 人気順
            End With ' OddsSanrenInfo
        Next i
        .TotalHyosuSanrenpuku = IncMid(bytBuf, p, 11)       '' 3連複票数合計
        .crlf = IncMid(bytBuf, p, 2)        '' レコード区切り
    End With

    'バッファ領域解放
    Erase bytBuf
   
    End Sub


    '****** １１．競走馬マスタ ****************************************

    Public Sub SetData_UM(ByVal lBuf As String, ByRef mBuf As JV_UM_UMA)
    Dim bytBuf() As Byte                            '' Byte列で処理するためのバッファ
    
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)      '' レコード種別
            .DataKubun = IncMid(bytBuf, p, 1)       '' データ区分
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)        '' 年
                .Month = IncMid(bytBuf, p, 2)       '' 月
                .Day = IncMid(bytBuf, p, 2)         '' 日
            End With ' MakeDate
        End With ' head
        .KettoNum = IncMid(bytBuf, p, 10)           '' 血統登録番号
        .DelKubun = IncMid(bytBuf, p, 1)            '' 競走馬抹消区分
        With .RegDate
            .Year = IncMid(bytBuf, p, 4)            '' 年
            .Month = IncMid(bytBuf, p, 2)           '' 月
            .Day = IncMid(bytBuf, p, 2)             '' 日
        End With ' RegDate
        With .DelDate
            .Year = IncMid(bytBuf, p, 4)            '' 年
            .Month = IncMid(bytBuf, p, 2)           '' 月
            .Day = IncMid(bytBuf, p, 2)             '' 日
        End With ' DelDate
        With .BirthDate
            .Year = IncMid(bytBuf, p, 4)            '' 年
            .Month = IncMid(bytBuf, p, 2)           '' 月
            .Day = IncMid(bytBuf, p, 2)             '' 日
        End With ' BirthDate
        .Bamei = IncMid(bytBuf, p, 36)              '' 馬名
        .BameiKana = IncMid(bytBuf, p, 36)          '' 馬名半角カナ
        .BameiEng = IncMid(bytBuf, p, 80)           '' 馬名欧字
        .UmaKigoCD = IncMid(bytBuf, p, 2)           '' 馬記号コード
        .SexCD = IncMid(bytBuf, p, 1)               '' 性別コード
        .HinsyuCD = IncMid(bytBuf, p, 1)            '' 品種コード
        .KeiroCD = IncMid(bytBuf, p, 2)             '' 毛色コード
        For i = 0 To 13
            With .Ketto3Info(i)
                .HansyokuNum = IncMid(bytBuf, p, 8) '' 繁殖登録番号
                .Bamei = IncMid(bytBuf, p, 36)      '' 馬名
            End With ' Ketto3Info
        Next i
        .TozaiCD = IncMid(bytBuf, p, 1)             '' 東西所属コード
        .ChokyosiCode = IncMid(bytBuf, p, 5)        '' 調教師コード
        .ChokyosiRyakusyo = IncMid(bytBuf, p, 8)    '' 調教師名略称
        .Syotai = IncMid(bytBuf, p, 20)             '' 招待地域名
        .BreederCode = IncMid(bytBuf, p, 6)         '' 生産者コード
        .BreederName = IncMid(bytBuf, p, 70)        '' 生産者名
        .SanchiName = IncMid(bytBuf, p, 20)         '' 産地名
        .BanusiCode = IncMid(bytBuf, p, 6)          '' 馬主コード
        .BanusiName = IncMid(bytBuf, p, 64)         '' 馬主名
        .RuikeiHonsyoHeiti = IncMid(bytBuf, p, 9)   '' 平地本賞金累計
        .RuikeiHonsyoSyogai = IncMid(bytBuf, p, 9)  '' 障害本賞金累計
        .RuikeiFukaHeichi = IncMid(bytBuf, p, 9)    '' 平地付加賞金累計
        .RuikeiFukaSyogai = IncMid(bytBuf, p, 9)    '' 障害付加賞金累計
        .RuikeiSyutokuHeichi = IncMid(bytBuf, p, 9) '' 平地収得賞金累計
        .RuikeiSyutokuSyogai = IncMid(bytBuf, p, 9) '' 障害収得賞金累計
        With .ChakuSogo
            For j = 0 To 5
                .Chakukaisu(j) = IncMid(bytBuf, p, 3)
            Next j
        End With ' ChakuSogo
        With .ChakuChuo
            For j = 0 To 5
                .Chakukaisu(j) = IncMid(bytBuf, p, 3)
            Next j
        End With ' ChakuChuo
        For i = 0 To 6
            With .ChakuKaisuBa(i)
                For j = 0 To 5
                    .Chakukaisu(j) = IncMid(bytBuf, p, 3)
                Next j
            End With ' ChakuKaisuBa
        Next i
        For i = 0 To 11
            With .ChakuKaisuJyotai(i)
                For j = 0 To 5
                    .Chakukaisu(j) = IncMid(bytBuf, p, 3)
                Next j
            End With ' ChakuKaisuJyotai
        Next i
        For i = 0 To 5
            With .ChakuKaisuKyori(i)
                For j = 0 To 5
                    .Chakukaisu(j) = IncMid(bytBuf, p, 3)
                Next j
            End With ' ChakuKaisuKyoriu
        Next i
        For i = 0 To 3
            .Kyakusitu(i) = IncMid(bytBuf, p, 3)    '' 脚質傾向
        Next i
        .RaceCount = IncMid(bytBuf, p, 3)           '' 登録レース数
        .crlf = IncMid(bytBuf, p, 2)                '' レコード区切り
    End With

    'バッファ領域解放
    Erase bytBuf
   
    End Sub


    '****** １２．騎手マスタ ****************************************

    Public Sub SetData_KS(lBuf As String, ByRef mBuf As JV_KS_KISYU)
    Dim bytBuf() As Byte                            '' Byte列で処理するためのバッファ
    
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)      '' レコード種別
            .DataKubun = IncMid(bytBuf, p, 1)       '' データ区分
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)        '' 年
                .Month = IncMid(bytBuf, p, 2)       '' 月
                .Day = IncMid(bytBuf, p, 2) '' 日
            End With ' MakeDate
        End With ' head
        .KisyuCode = IncMid(bytBuf, p, 5)   '' 騎手コード
        .DelKubun = IncMid(bytBuf, p, 1)    '' 騎手抹消区分
        With .IssueDate
            .Year = IncMid(bytBuf, p, 4)    '' 年
            .Month = IncMid(bytBuf, p, 2)   '' 月
            .Day = IncMid(bytBuf, p, 2)     '' 日
        End With ' IssueDate
        With .DelDate
            .Year = IncMid(bytBuf, p, 4)    '' 年
            .Month = IncMid(bytBuf, p, 2)   '' 月
            .Day = IncMid(bytBuf, p, 2)     '' 日
        End With ' DelDate
        With .BirthDate
            .Year = IncMid(bytBuf, p, 4)    '' 年
            .Month = IncMid(bytBuf, p, 2)   '' 月
            .Day = IncMid(bytBuf, p, 2)     '' 日
        End With ' BirthDate
        .KisyuName = IncMid(bytBuf, p, 34)  '' 騎手名漢字
        .reserved = IncMid(bytBuf, p, 34)   '' 予備
        .KisyuNameKana = IncMid(bytBuf, p, 30)      '' 騎手名半角カナ
        .KisyuRyakusyo = IncMid(bytBuf, p, 8)       '' 騎手名略称
        .KisyuNameEng = IncMid(bytBuf, p, 80)       '' 騎手名欧字
        .SexCD = IncMid(bytBuf, p, 1)       '' 性別区分
        .SikakuCD = IncMid(bytBuf, p, 1)    '' 騎乗資格コード
        .MinaraiCD = IncMid(bytBuf, p, 1)   '' 騎手見習コード
        .TozaiCD = IncMid(bytBuf, p, 1)     '' 騎手東西所属コード
        .Syotai = IncMid(bytBuf, p, 20)     '' 招待地域名
        .ChokyosiCode = IncMid(bytBuf, p, 5)        '' 所属調教師コード
        .ChokyosiRyakusyo = IncMid(bytBuf, p, 8)    '' 所属調教師名略称
        For i = 0 To 1
            With .HatuKiJyo(i)
                With .Hatukijyoid
                    .Year = IncMid(bytBuf, p, 4)    '' 開催年
                    .MonthDay = IncMid(bytBuf, p, 4)        '' 開催月日
                    .JyoCD = IncMid(bytBuf, p, 2)   '' 競馬場コード
                    .Kaiji = IncMid(bytBuf, p, 2)   '' 開催回[第N回]
                    .Nichiji = IncMid(bytBuf, p, 2) '' 開催日目[N日目]
                    .racenum = IncMid(bytBuf, p, 2) '' レース番号
                End With ' Hatukijyoid
                .SyussoTosu = IncMid(bytBuf, p, 2)  '' 出走頭数
                .KettoNum = IncMid(bytBuf, p, 10)   '' 血統登録番号
                .Bamei = IncMid(bytBuf, p, 36)      '' 馬名
                .KakuteiJyuni = IncMid(bytBuf, p, 2)        '' 確定着順
                .IJyoCD = IncMid(bytBuf, p, 1)      '' 異常区分コード
            End With ' HatuKiJyo
        Next i
        For i = 0 To 1
            With .HatuSyori(i)
                With .Hatusyoriid
                    .Year = IncMid(bytBuf, p, 4)    '' 開催年
                    .MonthDay = IncMid(bytBuf, p, 4)        '' 開催月日
                    .JyoCD = IncMid(bytBuf, p, 2)   '' 競馬場コード
                    .Kaiji = IncMid(bytBuf, p, 2)   '' 開催回[第N回]
                    .Nichiji = IncMid(bytBuf, p, 2) '' 開催日目[N日目]
                    .racenum = IncMid(bytBuf, p, 2) '' レース番号
                End With ' Hatusyoriid
                .SyussoTosu = IncMid(bytBuf, p, 2)  '' 出走頭数
                .KettoNum = IncMid(bytBuf, p, 10)   '' 血統登録番号
                .Bamei = IncMid(bytBuf, p, 36)      '' 馬名
            End With ' HatuSyori
        Next i
        For i = 0 To 2
            With .SaikinJyusyo(i)
                With .SaikinJyusyoid
                    .Year = IncMid(bytBuf, p, 4)    '' 開催年
                    .MonthDay = IncMid(bytBuf, p, 4)        '' 開催月日
                    .JyoCD = IncMid(bytBuf, p, 2)   '' 競馬場コード
                    .Kaiji = IncMid(bytBuf, p, 2)   '' 開催回[第N回]
                    .Nichiji = IncMid(bytBuf, p, 2) '' 開催日目[N日目]
                    .racenum = IncMid(bytBuf, p, 2) '' レース番号
                End With ' SaikinJyusyoid
                .Hondai = IncMid(bytBuf, p, 60)     '' 競走名本題
                .Ryakusyo10 = IncMid(bytBuf, p, 20) '' 競走名略称10字
                .Ryakusyo6 = IncMid(bytBuf, p, 12)  '' 競走名略称6字
                .Ryakusyo3 = IncMid(bytBuf, p, 6)   '' 競走名略称3字
                .GradeCD = IncMid(bytBuf, p, 1)     '' グレードコード
                .SyussoTosu = IncMid(bytBuf, p, 2)  '' 出走頭数
                .KettoNum = IncMid(bytBuf, p, 10)   '' 血統登録番号
                .Bamei = IncMid(bytBuf, p, 36)      '' 馬名
            End With ' SaikinJyusyo
        Next i
        For i = 0 To 2
            With .HonZenRuikei(i)
                .SetYear = IncMid(bytBuf, p, 4)     '' 設定年
                .HonSyokinHeichi = IncMid(bytBuf, p, 10)    '' 平地本賞金合計
                .HonSyokinSyogai = IncMid(bytBuf, p, 10)    '' 障害本賞金合計
                .FukaSyokinHeichi = IncMid(bytBuf, p, 10)   '' 平地付加賞金合計
                .FukaSyokinSyogai = IncMid(bytBuf, p, 10)   '' 障害付加賞金合計
                With .ChakuKaisuHeichi
                    For k = 0 To 5
                        .Chakukaisu(k) = IncMid(bytBuf, p, 6)
                    Next k
                End With ' ChakuKaisuHeichi
                With .ChakuKaisuSyogai
                    For k = 0 To 5
                        .Chakukaisu(k) = IncMid(bytBuf, p, 6)
                    Next k
                End With ' ChakuKaisuSyogai
                For j = 0 To 19
                    With .ChakuKaisuJyo(j)
                        For k = 0 To 5
                            .Chakukaisu(k) = IncMid(bytBuf, p, 6)
                        Next k
                    End With ' ChakuKaisuJyo
                Next j
                For j = 0 To 5
                    With .ChakuKaisuKyori(j)
                        For k = 0 To 5
                            .Chakukaisu(k) = IncMid(bytBuf, p, 6)
                        Next k
                    End With ' ChakuKaisuKyori
                Next j
            End With ' HonZenRuikei
        Next i
        .crlf = IncMid(bytBuf, p, 2)        '' レコード区切り
    End With

    'バッファ領域解放
    Erase bytBuf
    
    End Sub


    '****** １３．調教師マスタ ****************************************

    Public Sub SetData_CH(lBuf As String, ByRef mBuf As JV_CH_CHOKYOSI)
    Dim bytBuf() As Byte                            '' Byte列で処理するためのバッファ
    
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)              '' レコード種別
            .DataKubun = IncMid(bytBuf, p, 1)               '' データ区分
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)                '' 年
                .Month = IncMid(bytBuf, p, 2)               '' 月
                .Day = IncMid(bytBuf, p, 2)                 '' 日
            End With ' MakeDate
        End With ' head
        .ChokyosiCode = IncMid(bytBuf, p, 5)                '' 調教師コード
        .DelKubun = IncMid(bytBuf, p, 1)                    '' 調教師抹消区分
        With .IssueDate
            .Year = IncMid(bytBuf, p, 4)                    '' 年
            .Month = IncMid(bytBuf, p, 2)                   '' 月
            .Day = IncMid(bytBuf, p, 2)                     '' 日
        End With ' IssueDate
        With .DelDate
            .Year = IncMid(bytBuf, p, 4)                    '' 年
            .Month = IncMid(bytBuf, p, 2)                   '' 月
            .Day = IncMid(bytBuf, p, 2)                     '' 日
        End With ' DelDate
        With .BirthDate
            .Year = IncMid(bytBuf, p, 4)                    '' 年
            .Month = IncMid(bytBuf, p, 2)                   '' 月
            .Day = IncMid(bytBuf, p, 2)                     '' 日
        End With ' BirthDate
        .ChokyosiName = IncMid(bytBuf, p, 34)               '' 調教師名漢字
        .ChokyosiNameKana = IncMid(bytBuf, p, 30)           '' 調教師名半角カナ
        .ChokyosiRyakusyo = IncMid(bytBuf, p, 8)            '' 調教師名略称
        .ChokyosiNameEng = IncMid(bytBuf, p, 80)            '' 調教師名欧字
        .SexCD = IncMid(bytBuf, p, 1)                       '' 性別区分
        .TozaiCD = IncMid(bytBuf, p, 1)                     '' 調教師東西所属コード
        .Syotai = IncMid(bytBuf, p, 20)                     '' 招待地域名
        For i = 0 To 2
            With .SaikinJyusyo(i)
                With .SaikinJyusyoid
                    .Year = IncMid(bytBuf, p, 4)            '' 開催年
                    .MonthDay = IncMid(bytBuf, p, 4)        '' 開催月日
                    .JyoCD = IncMid(bytBuf, p, 2)           '' 競馬場コード
                    .Kaiji = IncMid(bytBuf, p, 2)           '' 開催回[第N回]
                    .Nichiji = IncMid(bytBuf, p, 2)         '' 開催日目[N日目]
                    .racenum = IncMid(bytBuf, p, 2)         '' レース番号
                End With ' SaikinJyusyoid
                .Hondai = IncMid(bytBuf, p, 60)             '' 競走名本題
                .Ryakusyo10 = IncMid(bytBuf, p, 20)         '' 競走名略称10字
                .Ryakusyo6 = IncMid(bytBuf, p, 12)          '' 競走名略称6字
                .Ryakusyo3 = IncMid(bytBuf, p, 6)           '' 競走名略称3字
                .GradeCD = IncMid(bytBuf, p, 1)             '' グレードコード
                .SyussoTosu = IncMid(bytBuf, p, 2)          '' 出走頭数
                .KettoNum = IncMid(bytBuf, p, 10)           '' 血統登録番号
                .Bamei = IncMid(bytBuf, p, 36)              '' 馬名
            End With ' SaikinJyusyo
        Next i
        For i = 0 To 2
            With .HonZenRuikei(i)
                .SetYear = IncMid(bytBuf, p, 4)             '' 設定年
                .HonSyokinHeichi = IncMid(bytBuf, p, 10)    '' 平地本賞金合計
                .HonSyokinSyogai = IncMid(bytBuf, p, 10)    '' 障害本賞金合計
                .FukaSyokinHeichi = IncMid(bytBuf, p, 10)   '' 平地付加賞金合計
                .FukaSyokinSyogai = IncMid(bytBuf, p, 10)   '' 障害付加賞金合計
                With .ChakuKaisuHeichi
                    For k = 0 To 5
                        .Chakukaisu(k) = IncMid(bytBuf, p, 6)
                    Next k
                End With ' ChakuKaisuHeichi
                With .ChakuKaisuSyogai
                    For k = 0 To 5
                        .Chakukaisu(k) = IncMid(bytBuf, p, 6)
                    Next k
                End With ' ChakuKaisuSyogai
                For j = 0 To 19
                    With .ChakuKaisuJyo(j)
                        For k = 0 To 5
                            .Chakukaisu(k) = IncMid(bytBuf, p, 6)
                        Next k
                    End With ' ChakuKaisuJyo
                Next j
                For j = 0 To 5
                    With .ChakuKaisuKyori(j)
                        For k = 0 To 5
                            .Chakukaisu(k) = IncMid(bytBuf, p, 6)
                        Next k
                    End With ' ChakuKaisuKyori
                Next j
            End With ' HonZenRuikei
        Next i
        .crlf = IncMid(bytBuf, p, 2)        '' レコード区切り
    End With

    'バッファ領域解放
    Erase bytBuf
   
    End Sub


    '******１４．生産者マスタ ****************************************

    Public Sub SetData_BR(lBuf As String, ByRef mBuf As JV_BR_BREEDER)
    Dim bytBuf() As Byte                            '' Byte列で処理するためのバッファ
    
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)              '' レコード種別
            .DataKubun = IncMid(bytBuf, p, 1)               '' データ区分
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)                '' 年
                .Month = IncMid(bytBuf, p, 2)               '' 月
                .Day = IncMid(bytBuf, p, 2)                 '' 日
            End With ' MakeDate
        End With ' head
        .BreederCode = IncMid(bytBuf, p, 6)                 '' 生産者コード
        .BreederName_Co = IncMid(bytBuf, p, 70)             '' 生産者名(法人格有）
        .BreederName = IncMid(bytBuf, p, 70)                '' 生産者名(法人格無）
        .BreederNameKana = IncMid(bytBuf, p, 70)            '' 生産者名半角カナ
        .BreederNameEng = IncMid(bytBuf, p, 168)            '' 生産者名欧字
        .Address = IncMid(bytBuf, p, 20)                    '' 生産者住所自治省名
        For i = 0 To 1
            With .HonRuikei(i)
                .SetYear = IncMid(bytBuf, p, 4)             '' 設定年
                .HonSyokinTotal = IncMid(bytBuf, p, 10)     '' 本賞金合計
                .Fukasyokin = IncMid(bytBuf, p, 10)         '' 付加賞金合計
                For j = 0 To 5
                    .Chakukaisu(j) = IncMid(bytBuf, p, 6)   '' 着回数
                Next j
            End With ' HonRuikei
        Next i
        .crlf = IncMid(bytBuf, p, 2)                        '' レコード区切り
    End With

    'バッファ領域解放
    Erase bytBuf
    
    End Sub


    '****** １５．馬主マスタ ****************************************

    Public Sub SetData_BN(lBuf As String, ByRef mBuf As JV_BN_BANUSI)
    Dim bytBuf() As Byte                            '' Byte列で処理するためのバッファ
    
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)              '' レコード種別
            .DataKubun = IncMid(bytBuf, p, 1)               '' データ区分
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)                '' 年
                .Month = IncMid(bytBuf, p, 2)               '' 月
                .Day = IncMid(bytBuf, p, 2)                 '' 日
            End With ' MakeDate
        End With ' head
        .BanusiCode = IncMid(bytBuf, p, 6)                  '' 馬主コード
        .BanusiName_Co = IncMid(bytBuf, p, 64)              '' 馬主名（法人格有）
        .BanusiName = IncMid(bytBuf, p, 64)                 '' 馬主名（法人格無）
        .BanusiNameKana = IncMid(bytBuf, p, 50)             '' 馬主名半角カナ
        .BanusiNameEng = IncMid(bytBuf, p, 100)             '' 馬主名欧字
        .Fukusyoku = IncMid(bytBuf, p, 60)                  '' 服色標示
        For i = 0 To 1
            With .HonRuikei(i)
                .SetYear = IncMid(bytBuf, p, 4)             '' 設定年
                .HonSyokinTotal = IncMid(bytBuf, p, 10)     '' 本賞金合計
                .Fukasyokin = IncMid(bytBuf, p, 10)         '' 付加賞金合計
                For j = 0 To 5
                    .Chakukaisu(j) = IncMid(bytBuf, p, 6)   '' 着回数
                Next j
            End With ' HonRuikei
        Next i
        .crlf = IncMid(bytBuf, p, 2)                        '' レコード区切り
    End With

    'バッファ領域解放
    Erase bytBuf
   
    End Sub

    '****** １６．繁殖馬マスタ ****************************************

    Public Sub SetData_HN(lBuf As String, ByRef mBuf As JV_HN_HANSYOKU)
    Dim bytBuf() As Byte                            '' Byte列で処理するためのバッファ
    
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)      '' レコード種別
            .DataKubun = IncMid(bytBuf, p, 1)       '' データ区分
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)        '' 年
                .Month = IncMid(bytBuf, p, 2)       '' 月
                .Day = IncMid(bytBuf, p, 2)         '' 日
            End With ' MakeDate
        End With ' head
        .HansyokuNum = IncMid(bytBuf, p, 8)         '' 繁殖登録番号
        .reserved = IncMid(bytBuf, p, 8)            '' 予備
        .KettoNum = IncMid(bytBuf, p, 10)           '' 血統登録番号
        .DelKubun = IncMid(bytBuf, p, 1)            '' 繁殖馬抹消区分
        .Bamei = IncMid(bytBuf, p, 36)              '' 馬名
        .BameiKana = IncMid(bytBuf, p, 40)          '' 馬名半角カナ
        .BameiEng = IncMid(bytBuf, p, 80)           '' 馬名欧字
        .BirthYear = IncMid(bytBuf, p, 4)           '' 生年
        .SexCD = IncMid(bytBuf, p, 1)               '' 性別コード
        .HinsyuCD = IncMid(bytBuf, p, 1)            '' 品種コード
        .KeiroCD = IncMid(bytBuf, p, 2)             '' 毛色コード
        .HansyokuMochiKubun = IncMid(bytBuf, p, 1)  '' 繁殖馬持込区分
        .ImportYear = IncMid(bytBuf, p, 4)          '' 輸入年
        .SanchiName = IncMid(bytBuf, p, 20)         '' 産地名
        .HansyokuFNum = IncMid(bytBuf, p, 8)        '' 父馬繁殖登録番号
        .HansyokuMNum = IncMid(bytBuf, p, 8)        '' 母馬繁殖登録番号
        .crlf = IncMid(bytBuf, p, 2)                '' レコード区切り
    End With

    'バッファ領域解放
    Erase bytBuf
   
    End Sub


    '****** １７．産駒マスタ ****************************************

    Public Sub SetData_SK(lBuf As String, ByRef mBuf As JV_SK_SANKU)
    Dim bytBuf() As Byte                            '' Byte列で処理するためのバッファ
    
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)      '' レコード種別
            .DataKubun = IncMid(bytBuf, p, 1)       '' データ区分
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)        '' 年
                .Month = IncMid(bytBuf, p, 2)       '' 月
                .Day = IncMid(bytBuf, p, 2)         '' 日
            End With ' MakeDate
        End With ' head
        .KettoNum = IncMid(bytBuf, p, 10)           '' 血統登録番号
        With .BirthDate
            .Year = IncMid(bytBuf, p, 4)            '' 年
            .Month = IncMid(bytBuf, p, 2)           '' 月
            .Day = IncMid(bytBuf, p, 2)             '' 日
        End With ' BirthDate
        .SexCD = IncMid(bytBuf, p, 1)               '' 性別コード
        .HinsyuCD = IncMid(bytBuf, p, 1)            '' 品種コード
        .KeiroCD = IncMid(bytBuf, p, 2)             '' 毛色コード
        .SankuMochiKubun = IncMid(bytBuf, p, 1)     '' 産駒持込区分
        .ImportYear = IncMid(bytBuf, p, 4)          '' 輸入年
        .BreederCode = IncMid(bytBuf, p, 6)         '' 生産者コード
        .SanchiName = IncMid(bytBuf, p, 20)         '' 産地名
        For i = 0 To 13
            .HansyokuNum(i) = IncMid(bytBuf, p, 8)  '' 3代血統
        Next i
    End With

    'バッファ領域解放
    Erase bytBuf
    
    End Sub

    '****** １８．レコードマスタ ****************************************

    Public Sub SetData_RC(lBuf As String, ByRef mBuf As JV_RC_RECORD)
    Dim bytBuf() As Byte                            '' Byte列で処理するためのバッファ
    
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)              '' レコード種別
            .DataKubun = IncMid(bytBuf, p, 1)               '' データ区分
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)                '' 年
                .Month = IncMid(bytBuf, p, 2)               '' 月
                .Day = IncMid(bytBuf, p, 2)                 '' 日
            End With ' MakeDate
        End With ' head
        .RecInfoKubun = IncMid(bytBuf, p, 1)                '' レコード識別区分
        With .id
            .Year = IncMid(bytBuf, p, 4)                    '' 開催年
            .MonthDay = IncMid(bytBuf, p, 4)                '' 開催月日
            .JyoCD = IncMid(bytBuf, p, 2)                   '' 競馬場コード
            .Kaiji = IncMid(bytBuf, p, 2)                   '' 開催回[第N回]
            .Nichiji = IncMid(bytBuf, p, 2)                 '' 開催日目[N日目]
            .racenum = IncMid(bytBuf, p, 2)                 '' レース番号
        End With ' id
        .TokuNum = IncMid(bytBuf, p, 4)                     '' 特別競走番号
        .Hondai = IncMid(bytBuf, p, 60)                     '' 競走名本題
        .GradeCD = IncMid(bytBuf, p, 1)                     '' グレードコード
        .SyubetuCD = IncMid(bytBuf, p, 2)                   '' 競走種別コード
        .Kyori = IncMid(bytBuf, p, 4)                       '' 距離
        .TrackCD = IncMid(bytBuf, p, 2)                     '' トラックコード
        .RecKubun = IncMid(bytBuf, p, 1)                    '' レコード区分
        .RecTime = IncMid(bytBuf, p, 4)                     '' レコードタイム
        With .TenkoBaba
            .TenkoCD = IncMid(bytBuf, p, 1)                 '' 天候コード
            .SibaBabaCD = IncMid(bytBuf, p, 1)              '' 芝馬場状態コード
            .DirtBabaCD = IncMid(bytBuf, p, 1)              '' ダート馬場状態コード
        End With ' TenkoBaba
        For i = 0 To 2
            With .RecUmaInfo(i)
                .KettoNum = IncMid(bytBuf, p, 10)           '' 血統登録番号
                .Bamei = IncMid(bytBuf, p, 36)              '' 馬名
                .UmaKigoCD = IncMid(bytBuf, p, 2)           '' 馬記号コード
                .SexCD = IncMid(bytBuf, p, 1)               '' 性別コード
                .ChokyosiCode = IncMid(bytBuf, p, 5)        '' 調教師コード
                .ChokyosiName = IncMid(bytBuf, p, 34)       '' 調教師名
                .Futan = IncMid(bytBuf, p, 3)               '' 負担重量
                .KisyuCode = IncMid(bytBuf, p, 5)           '' 騎手コード
                .KisyuName = IncMid(bytBuf, p, 34)          '' 騎手名
            End With ' RecUmaInfo
        Next i
        .crlf = IncMid(bytBuf, p, 2)                        '' レコード区切り
    End With

    'バッファ領域解放
    Erase bytBuf

    End Sub


    '****** １９．坂路調教 ****************************************

    Public Sub SetData_HC(lBuf As String, ByRef mBuf As JV_HC_HANRO)
    Dim bytBuf() As Byte                            '' Byte列で処理するためのバッファ
    
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)  '' レコード種別
            .DataKubun = IncMid(bytBuf, p, 1)   '' データ区分
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)    '' 年
                .Month = IncMid(bytBuf, p, 2)   '' 月
                .Day = IncMid(bytBuf, p, 2)     '' 日
            End With ' MakeDate
        End With ' head
        .TresenKubun = IncMid(bytBuf, p, 1)     '' トレセン区分
        With .ChokyoDate
            .Year = IncMid(bytBuf, p, 4)        '' 年
            .Month = IncMid(bytBuf, p, 2)       '' 月
            .Day = IncMid(bytBuf, p, 2)         '' 日
        End With ' ChokyoDate
        .ChokyoTime = IncMid(bytBuf, p, 4)      '' 調教時刻
        .KettoNum = IncMid(bytBuf, p, 10)       '' 血統登録番号
        .HaronTime4 = IncMid(bytBuf, p, 4)      '' 4ハロンタイム合計(800M-0M)
        .LapTime4 = IncMid(bytBuf, p, 3)        '' ラップタイム(800M-600M)
        .HaronTime3 = IncMid(bytBuf, p, 4)      '' 3ハロンタイム合計(600M-0M)
        .LapTime3 = IncMid(bytBuf, p, 3)        '' ラップタイム(600M-400M)
        .HaronTime2 = IncMid(bytBuf, p, 4)      '' 2ハロンタイム合計(400M-0M)
        .LapTime2 = IncMid(bytBuf, p, 3)        '' ラップタイム(400M-200M)
        .LapTime1 = IncMid(bytBuf, p, 3)        '' ラップタイム(200M-0M)
        .crlf = IncMid(bytBuf, p, 2)            '' レコード区切り
    End With

    'バッファ領域解放
    Erase bytBuf
    
  End Sub


    '****** ２０．馬体重 ****************************************

    Public Sub SetData_WH(lBuf As String, ByRef mBuf As JV_WH_BATAIJYU)
    Dim bytBuf() As Byte                            '' Byte列で処理するためのバッファ
    
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)      '' レコード種別
            .DataKubun = IncMid(bytBuf, p, 1)       '' データ区分
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)        '' 年
                .Month = IncMid(bytBuf, p, 2)       '' 月
                .Day = IncMid(bytBuf, p, 2)         '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)            '' 開催年
            .MonthDay = IncMid(bytBuf, p, 4)        '' 開催月日
            .JyoCD = IncMid(bytBuf, p, 2)           '' 競馬場コード
            .Kaiji = IncMid(bytBuf, p, 2)           '' 開催回[第N回]
            .Nichiji = IncMid(bytBuf, p, 2)         '' 開催日目[N日目]
            .racenum = IncMid(bytBuf, p, 2)         '' レース番号
        End With ' id
        With .HappyoTime
            .Month = IncMid(bytBuf, p, 2)           '' 月
            .Day = IncMid(bytBuf, p, 2)             '' 日
            .Hour = IncMid(bytBuf, p, 2)            '' 時
            .Minute = IncMid(bytBuf, p, 2)          '' 分
        End With ' HappyoTime
        For i = 0 To 17
            With .BataijyuInfo(i)
                .Umaban = IncMid(bytBuf, p, 2)      '' 馬番
                .Bamei = IncMid(bytBuf, p, 36)      '' 馬名
                .BaTaijyu = IncMid(bytBuf, p, 3)    '' 馬体重
                .ZogenFugo = IncMid(bytBuf, p, 1)   '' 増減符号
                .ZogenSa = IncMid(bytBuf, p, 3)     '' 増減差
            End With ' BataijyuInfo
        Next i
        .crlf = IncMid(bytBuf, p, 2)                '' レコード区切り
    End With

    'バッファ領域解放
    Erase bytBuf
   
    End Sub


    '****** ２１．天候馬場状態 ******************************************

    Public Sub SetData_WE(lBuf As String, ByRef mBuf As JV_WE_WEATHER)
    Dim bytBuf() As Byte                            '' Byte列で処理するためのバッファ
    
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)      '' レコード種別
            .DataKubun = IncMid(bytBuf, p, 1)       '' データ区分
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)        '' 年
                .Month = IncMid(bytBuf, p, 2)       '' 月
                .Day = IncMid(bytBuf, p, 2)         '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)            '' 開催年
            .MonthDay = IncMid(bytBuf, p, 4)        '' 開催月日
            .JyoCD = IncMid(bytBuf, p, 2)           '' 競馬場コード
            .Kaiji = IncMid(bytBuf, p, 2)           '' 開催回[第N回]
            .Nichiji = IncMid(bytBuf, p, 2)         '' 開催日目[N日目]
        End With ' id
        With .HappyoTime
            .Month = IncMid(bytBuf, p, 2)           '' 月
            .Day = IncMid(bytBuf, p, 2)             '' 日
            .Hour = IncMid(bytBuf, p, 2)            '' 時
            .Minute = IncMid(bytBuf, p, 2)          '' 分
        End With ' HappyoTime
        .HenkoID = IncMid(bytBuf, p, 1)             '' 変更識別
        With .TenkoBaba
            .TenkoCD = IncMid(bytBuf, p, 1)         '' 天候コード
            .SibaBabaCD = IncMid(bytBuf, p, 1)      '' 芝馬場状態コード
            .DirtBabaCD = IncMid(bytBuf, p, 1)      '' ダート馬場状態コード
        End With ' TenkoBaba
        With .TenkoBabaBefore
            .TenkoCD = IncMid(bytBuf, p, 1)         '' 天候コード
            .SibaBabaCD = IncMid(bytBuf, p, 1)      '' 芝馬場状態コード
            .DirtBabaCD = IncMid(bytBuf, p, 1)      '' ダート馬場状態コード
        End With ' TenkoBabaBefore
        .crlf = IncMid(bytBuf, p, 2)                '' レコード区切り
    End With

    'バッファ領域解放
    Erase bytBuf
    
    End Sub


    '****** ２２．出走取消・競争除外 ****************************************

    Public Sub SetData_AV(lBuf As String, ByRef mBuf As JV_AV_INFO)
    Dim bytBuf() As Byte                            '' Byte列で処理するためのバッファ
    
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)      '' レコード種別
            .DataKubun = IncMid(bytBuf, p, 1)       '' データ区分
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)        '' 年
                .Month = IncMid(bytBuf, p, 2)       '' 月
                .Day = IncMid(bytBuf, p, 2)         '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)            '' 開催年
            .MonthDay = IncMid(bytBuf, p, 4)        '' 開催月日
            .JyoCD = IncMid(bytBuf, p, 2)           '' 競馬場コード
            .Kaiji = IncMid(bytBuf, p, 2)           '' 開催回[第N回]
            .Nichiji = IncMid(bytBuf, p, 2)         '' 開催日目[N日目]
            .racenum = IncMid(bytBuf, p, 2)         '' レース番号
        End With ' id
        With .HappyoTime
            .Month = IncMid(bytBuf, p, 2)           '' 月
            .Day = IncMid(bytBuf, p, 2)             '' 日
            .Hour = IncMid(bytBuf, p, 2)            '' 時
            .Minute = IncMid(bytBuf, p, 2)          '' 分
        End With ' HappyoTime
        .Umaban = IncMid(bytBuf, p, 2)              '' 馬番
        .Bamei = IncMid(bytBuf, p, 36)              '' 馬名
        .JiyuKubun = IncMid(bytBuf, p, 3)           '' 事由区分
        .crlf = IncMid(bytBuf, p, 2)                '' レコード区切り
    End With

    'バッファ領域解放
    Erase bytBuf
    
    End Sub

    '************ ２３．騎手変更 ****************************************
  
    Public Sub SetData_JC(lBuf As String, ByRef mBuf As JV_JC_INFO)
    Dim bytBuf() As Byte                            '' Byte列で処理するためのバッファ
    
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)  '' レコード種別
            .DataKubun = IncMid(bytBuf, p, 1)   '' データ区分
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)    '' 年
                .Month = IncMid(bytBuf, p, 2)   '' 月
                .Day = IncMid(bytBuf, p, 2)     '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)        '' 開催年
            .MonthDay = IncMid(bytBuf, p, 4)        '' 開催月日
            .JyoCD = IncMid(bytBuf, p, 2)       '' 競馬場コード
            .Kaiji = IncMid(bytBuf, p, 2)       '' 開催回[第N回]
            .Nichiji = IncMid(bytBuf, p, 2)     '' 開催日目[N日目]
            .racenum = IncMid(bytBuf, p, 2)     '' レース番号
        End With ' id
        With .HappyoTime
            .Month = IncMid(bytBuf, p, 2)       '' 月
            .Day = IncMid(bytBuf, p, 2)         '' 日
            .Hour = IncMid(bytBuf, p, 2)        '' 時
            .Minute = IncMid(bytBuf, p, 2)      '' 分
        End With ' HappyoTime
        .Umaban = IncMid(bytBuf, p, 2)          '' 馬番
        .Bamei = IncMid(bytBuf, p, 36)          '' 馬名
        With .JCInfoAfter
            .Futan = IncMid(bytBuf, p, 3)       '' 負担重量
            .KisyuCode = IncMid(bytBuf, p, 5)   '' 騎手コード
            .KisyuName = IncMid(bytBuf, p, 34)  '' 騎手名
            .MinaraiCD = IncMid(bytBuf, p, 1)   '' 騎手見習コード
        End With ' JCInfoAfter
        With .JCInfoBefore
            .Futan = IncMid(bytBuf, p, 3)       '' 負担重量
            .KisyuCode = IncMid(bytBuf, p, 5)   '' 騎手コード
            .KisyuName = IncMid(bytBuf, p, 34)  '' 騎手名
            .MinaraiCD = IncMid(bytBuf, p, 1)   '' 騎手見習コード
        End With ' JCInfoBefore
        .crlf = IncMid(bytBuf, p, 2)            '' レコード区切り
    End With

    'バッファ領域解放
    Erase bytBuf
   
    End Sub

    '****** ２４．データマイニング予想***********************************
    
    Public Sub SetData_DM(lBuf As String, ByRef mBuf As JV_DM_INFO)
    Dim bytBuf() As Byte                            '' Byte列で処理するためのバッファ
    
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)  '' レコード種別
            .DataKubun = IncMid(bytBuf, p, 1)   '' データ区分
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)    '' 年
                .Month = IncMid(bytBuf, p, 2)   '' 月
                .Day = IncMid(bytBuf, p, 2)     '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)        '' 開催年
            .MonthDay = IncMid(bytBuf, p, 4)    '' 開催月日
            .JyoCD = IncMid(bytBuf, p, 2)       '' 競馬場コード
            .Kaiji = IncMid(bytBuf, p, 2)       '' 開催回[第N回]
            .Nichiji = IncMid(bytBuf, p, 2)     '' 開催日目[N日目]
            .racenum = IncMid(bytBuf, p, 2)     '' レース番号
        End With ' id
        With .MakeHM
            .Hour = IncMid(bytBuf, p, 2)        '' 時
            .Minute = IncMid(bytBuf, p, 2)      '' 分
        End With ' MakeHM
        For i = 0 To 17
            With .DMInfo(i)
                .Umaban = IncMid(bytBuf, p, 2)  '' 馬番
                .DMTime = IncMid(bytBuf, p, 5)  '' 予想走破タイム
                .DMGosaP = IncMid(bytBuf, p, 4) '' 予想誤差(信頼度)＋
                .DMGosaM = IncMid(bytBuf, p, 4) '' 予想誤差(信頼度)－
            End With ' DMInfo
        Next i
        .crlf = IncMid(bytBuf, p, 2)            '' レコード区切り
    End With

    'バッファ領域解放
    Erase bytBuf
    
    End Sub


    '****** ２５．開催スケジュール************************************
    
    Public Sub SetData_YS(lBuf As String, ByRef mBuf As JV_YS_SCHEDULE)
    Dim bytBuf() As Byte                            '' Byte列で処理するためのバッファ
    
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)      '' レコード種別
            .DataKubun = IncMid(bytBuf, p, 1)       '' データ区分
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)        '' 年
                .Month = IncMid(bytBuf, p, 2)       '' 月
                .Day = IncMid(bytBuf, p, 2)         '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)            '' 開催年
            .MonthDay = IncMid(bytBuf, p, 4)        '' 開催月日
            .JyoCD = IncMid(bytBuf, p, 2)           '' 競馬場コード
            .Kaiji = IncMid(bytBuf, p, 2)           '' 開催回[第N回]
            .Nichiji = IncMid(bytBuf, p, 2)         '' 開催日目[N日目]
        End With ' id
        .YoubiCD = IncMid(bytBuf, p, 1)             '' 曜日コード
        For i = 0 To 2
            With .JyusyoInfo(i)
                .TokuNum = IncMid(bytBuf, p, 4)     '' 特別競走番号
                .Hondai = IncMid(bytBuf, p, 60)     '' 競走名本題
                .Ryakusyo10 = IncMid(bytBuf, p, 20) '' 競走名略称10字
                .Ryakusyo6 = IncMid(bytBuf, p, 12)  '' 競走名略称6字
                .Ryakusyo3 = IncMid(bytBuf, p, 6)   '' 競走名略称3字
                .Nkai = IncMid(bytBuf, p, 3)        '' 重賞回次[第N回]
                .GradeCD = IncMid(bytBuf, p, 1)     '' グレードコード
                .SyubetuCD = IncMid(bytBuf, p, 2)   '' 競走種別コード
                .KigoCD = IncMid(bytBuf, p, 3)      '' 競走記号コード
                .JyuryoCD = IncMid(bytBuf, p, 1)    '' 重量種別コード
                .Kyori = IncMid(bytBuf, p, 4)       '' 距離
                .TrackCD = IncMid(bytBuf, p, 2)     '' トラックコード
            End With ' JyusyoInfo
        Next i
        .crlf = IncMid(bytBuf, p, 2)                '' レコード区切り
    End With

    'バッファ領域解放
    Erase bytBuf
    
    End Sub

    Public Sub SetData_H6(lBuf As String, ByRef mBuf As JV_H6_HYOSU_SANRENTAN)
    Dim bytBuf() As Byte                                    '' バイト配列で処理するためのバッファ
    Dim i As Integer                                        '' ループカウンタ
    Dim j As Integer                                        '' ループカウンタ
    Dim k As Integer                                        '' ループカウンタ
    Dim p As Long                                           '' 切り分け開始位置
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)              '' レコード種別
            .DataKubun = IncMid(bytBuf, p, 1)               '' データ区分
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)                '' 年
                .Month = IncMid(bytBuf, p, 2)               '' 月
                .Day = IncMid(bytBuf, p, 2)                 '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)                    '' 開催年
            .MonthDay = IncMid(bytBuf, p, 4)                '' 開催月日
            .JyoCD = IncMid(bytBuf, p, 2)                   '' 競馬場コード
            .Kaiji = IncMid(bytBuf, p, 2)                   '' 開催回[第N回]
            .Nichiji = IncMid(bytBuf, p, 2)                 '' 開催日目[N日目]
            .racenum = IncMid(bytBuf, p, 2)                 '' レース番号
        End With ' id
        .TorokuTosu = IncMid(bytBuf, p, 2)                  '' 登録頭数
        .SyussoTosu = IncMid(bytBuf, p, 2)                  '' 出走頭数
        .HatubaiFlag = IncMid(bytBuf, p, 1)                     '' 発売フラグ
        For i = 0 To 17
            .HenkanUma(i) = IncMid(bytBuf, p, 1)            '' 返還馬番情報(馬番01～18)
        Next i

        For i = 0 To 4895
            With .HyoSanrentan(i)
                .Kumi = IncMid(bytBuf, p, 6)                '' 組番
                .Hyo = IncMid(bytBuf, p, 11)                '' 票数
                .Ninki = IncMid(bytBuf, p, 4)               '' 人気
            End With ' HyoSanrentan
        Next i
        For i = 0 To 1
            .HyoTotal(i) = IncMid(bytBuf, p, 11)            '' 票数合計
        Next i
        .crlf = IncMid(bytBuf, p, 2)                        '' レコード区切り
    End With
    
    'バッファ領域解放
    Erase bytBuf
    
    End Sub

    Public Sub SetData_O6(lBuf As String, ByRef mBuf As JV_O6_ODDS_SANRENTAN)

    Dim bytBuf() As Byte                                    '' バイト配列で処理するためのバッファ
    Dim i As Integer                                        '' ループカウンタ
    Dim j As Integer                                        '' ループカウンタ
    Dim k As Integer                                        '' ループカウンタ
    Dim p As Long                                           '' 切り分け開始位置
    
    bytBuf = StrConv(lBuf, vbFromUnicode)

    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)              '' レコード種別
            .DataKubun = IncMid(bytBuf, p, 1)               '' データ区分
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)                '' 年
                .Month = IncMid(bytBuf, p, 2)               '' 月
                .Day = IncMid(bytBuf, p, 2)                 '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)                    '' 開催年
            .MonthDay = IncMid(bytBuf, p, 4)                '' 開催月日
            .JyoCD = IncMid(bytBuf, p, 2)                   '' 競馬場コード
            .Kaiji = IncMid(bytBuf, p, 2)                   '' 開催回[第N回]
            .Nichiji = IncMid(bytBuf, p, 2)                 '' 開催日目[N日目]
            .racenum = IncMid(bytBuf, p, 2)                 '' レース番号
        End With ' id
        With .HappyoTime
            .Month = IncMid(bytBuf, p, 2)                   '' 月
            .Day = IncMid(bytBuf, p, 2)                     '' 日
            .Hour = IncMid(bytBuf, p, 2)                    '' 時
            .Minute = IncMid(bytBuf, p, 2)                  '' 分
        End With ' HappyoTime
        .TorokuTosu = IncMid(bytBuf, p, 2)                  '' 登録頭数
        .SyussoTosu = IncMid(bytBuf, p, 2)                  '' 出走頭数
        .SanrentanFlag = IncMid(bytBuf, p, 1)               '' 発売フラグ　3連単
        For i = 0 To 4895
            With .OddsSanrentanInfo(i)
                .Kumi = IncMid(bytBuf, p, 6)                '' 組番
                .Odds = IncMid(bytBuf, p, 7)                '' オッズ
                .Ninki = IncMid(bytBuf, p, 4)               '' 人気順
            End With ' OddsSanrentanInfo
        Next i
        .TotalHyosuSanrentan = IncMid(bytBuf, p, 11)        '' 3連単票数合計
        .crlf = IncMid(bytBuf, p, 2)                        '' レコード区切り
    End With
    
    'バッファ領域解放
    Erase bytBuf
    
    End Sub

Public Sub SetData_O6Z(lBuf As String, ByRef mBuf As JV_O6_ODDS_SANRENTAN2)

    Dim bytBuf() As Byte                                    '' バイト配列で処理するためのバッファ
    Dim i As Integer                                        '' ループカウンタ
    Dim j As Integer                                        '' ループカウンタ
    Dim k As Integer                                        '' ループカウンタ
    Dim p As Long                                           '' 切り分け開始位置
    
    bytBuf = StrConv(lBuf, vbFromUnicode)

    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)              '' レコード種別
            .DataKubun = IncMid(bytBuf, p, 1)               '' データ区分
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)                '' 年
                .Month = IncMid(bytBuf, p, 2)               '' 月
                .Day = IncMid(bytBuf, p, 2)                 '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)                    '' 開催年
            .MonthDay = IncMid(bytBuf, p, 4)                '' 開催月日
            .JyoCD = IncMid(bytBuf, p, 2)                   '' 競馬場コード
            .Kaiji = IncMid(bytBuf, p, 2)                   '' 開催回[第N回]
            .Nichiji = IncMid(bytBuf, p, 2)                 '' 開催日目[N日目]
            .racenum = IncMid(bytBuf, p, 2)                 '' レース番号
        End With ' id
        With .HappyoTime
            .Month = IncMid(bytBuf, p, 2)                   '' 月
            .Day = IncMid(bytBuf, p, 2)                     '' 日
            .Hour = IncMid(bytBuf, p, 2)                    '' 時
            .Minute = IncMid(bytBuf, p, 2)                  '' 分
        End With ' HappyoTime
        .TorokuTosu = IncMid(bytBuf, p, 2)                  '' 登録頭数
        .SyussoTosu = IncMid(bytBuf, p, 2)                  '' 出走頭数
        .SanrentanFlag = IncMid(bytBuf, p, 1)               '' 発売フラグ　3連単
        Set .OddsSanrentanInfo = New Collection
        For i = 0 To 4895
            Set cOddssanrentaninfo = New cODDS_SANRENTAN_INFO
            cOddssanrentaninfo.Kumi = IncMid(bytBuf, p, 6)
            cOddssanrentaninfo.Odds = IncMid(bytBuf, p, 7)
            cOddssanrentaninfo.Ninki = IncMid(bytBuf, p, 4)
            .OddsSanrentanInfo.Add cOddssanrentaninfo
        Next i
'        Debug.Print .OddsSanrentanInfo.Count
        .TotalHyosuSanrentan = IncMid(bytBuf, p, 11)        '' 3連単票数合計
        .crlf = IncMid(bytBuf, p, 2)                        '' レコード区切り
    End With
    
    'バッファ領域解放
    Erase bytBuf
    
    End Sub

    Public Sub SetData_CC(lBuf As String, ByRef mBuf As JV_CC_INFO)

    Dim bytBuf() As Byte                                    '' バイト配列で処理するためのバッファ
    Dim i As Integer                                        '' ループカウンタ
    Dim j As Integer                                        '' ループカウンタ
    Dim k As Integer                                        '' ループカウンタ
    Dim p As Long                                           '' 切り分け開始位置
    
    bytBuf = StrConv(lBuf, vbFromUnicode)

    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)              '' レコード種別
            .DataKubun = IncMid(bytBuf, p, 1)               '' データ区分
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)                '' 年
                .Month = IncMid(bytBuf, p, 2)               '' 月
                .Day = IncMid(bytBuf, p, 2)                 '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)                    '' 開催年
            .MonthDay = IncMid(bytBuf, p, 4)                    '' 開催月日
            .JyoCD = IncMid(bytBuf, p, 2)                   '' 競馬場コード
            .Kaiji = IncMid(bytBuf, p, 2)                   '' 開催回[第N回]
            .Nichiji = IncMid(bytBuf, p, 2)                 '' 開催日目[N日目]
            .racenum = IncMid(bytBuf, p, 2)                 '' レース番号
        End With ' id
        With .HappyoTime
            .Month = IncMid(bytBuf, p, 2)                   '' 月
            .Day = IncMid(bytBuf, p, 2)                     '' 日
            .Hour = IncMid(bytBuf, p, 2)                    '' 時
            .Minute = IncMid(bytBuf, p, 2)                  '' 分
        End With ' HappyoTime
        
        With .CCInfoAfter
            .Kyori = IncMid(bytBuf, p, 4)                   '' 距離
            .TruckCD = IncMid(bytBuf, p, 2)                 '' トラックコード
        End With ' CCInfoAfter
        With .CCInfoBefore
            .Kyori = IncMid(bytBuf, p, 4)                   '' 距離
            .TruckCD = IncMid(bytBuf, p, 2)                 '' トラックコード
        End With ' CCInfoBefore
        .JiyuCD = IncMid(bytBuf, p, 1)                      '' 事由コード

        .crlf = IncMid(bytBuf, p, 2)                        '' レコード区切り
    End With
    
    'バッファ領域解放
    Erase bytBuf
    
    End Sub

    Public Sub SetData_TC(lBuf As String, ByRef mBuf As JV_TC_INFO)
    Dim bytBuf() As Byte                                    '' バイト配列で処理するためのバッファ
    Dim i As Integer                                        '' ループカウンタ
    Dim j As Integer                                        '' ループカウンタ
    Dim k As Integer                                        '' ループカウンタ
    Dim p As Long                                           '' 切り分け開始位置
    
    bytBuf = StrConv(lBuf, vbFromUnicode)

    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)              '' レコード種別
            .DataKubun = IncMid(bytBuf, p, 1)               '' データ区分
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)                '' 年
                .Month = IncMid(bytBuf, p, 2)               '' 月
                .Day = IncMid(bytBuf, p, 2)                 '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)                    '' 開催年
            .MonthDay = IncMid(bytBuf, p, 4)                    '' 開催月日
            .JyoCD = IncMid(bytBuf, p, 2)                   '' 競馬場コード
            .Kaiji = IncMid(bytBuf, p, 2)                   '' 開催回[第N回]
            .Nichiji = IncMid(bytBuf, p, 2)                 '' 開催日目[N日目]
            .racenum = IncMid(bytBuf, p, 2)                 '' レース番号
        End With ' id
        With .HappyoTime
            .Month = IncMid(bytBuf, p, 2)                   '' 月
            .Day = IncMid(bytBuf, p, 2)                     '' 日
            .Hour = IncMid(bytBuf, p, 2)                    '' 時
            .Minute = IncMid(bytBuf, p, 2)                  '' 分
        End With ' HappyoTime
        With .TCInfoAfter
            .Ji = IncMid(bytBuf, p, 2)                          '' 時
            .Fun = IncMid(bytBuf, p, 2)                         '' 分
        End With ' TCInfoAfter
        With .TCInfoBefore
            .Ji = IncMid(bytBuf, p, 2)                          '' 時
            .Fun = IncMid(bytBuf, p, 2)                         '' 分
        End With ' TCInfoBefore

        .crlf = IncMid(bytBuf, p, 2)                        '' レコード区切り
    End With
    
    'バッファ領域解放
    Erase bytBuf
    
    End Sub

 '------------------------------------------------------------------------
 '　　バイト配列をバイト長で切出し
 '------------------------------------------------------------------------
 Public Function IncMid(ByRef vBuf() As Byte, p As Long, length As Long) As String
     IncMid = StrConv(MidB(vBuf, p, length), vbUnicode)
     p = p + length
 End Function
     
        
