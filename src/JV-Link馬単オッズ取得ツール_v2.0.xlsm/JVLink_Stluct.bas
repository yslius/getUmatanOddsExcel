Attribute VB_Name = "JVLink_Stluct"
Option Base 0


'  JRA-VAN Data Lab. JV-Data�\����
'
'
'   �쐬: JRA-VAN �\�t�g�E�F�A�H�[
'
'========================================================================
'   (C) Copyright Turf Media System Co.,Ltd. 2003 All rights reserved
'========================================================================


    '''''''''''''''''''' ���ʍ\���� ''''''''''''''''''''

 '<�N����>
 Private Type YMD
     Year   As String                    ''�N
     Month  As String                    ''��
     Day    As String                    ''��
 End Type


 '<�����b>
 Private Type HMS
     Hour   As String                    ''��
     Minute As String                    ''��
     Second As String                    ''�b
 End Type


 '<����>
 Private Type HM
     Hour As String                      ''��
     Minute As String                    ''��
 End Type


 '<��������>
 Private Type MDHM
     Month As String                     ''��
     Day As String                       ''��
     Hour As String                      ''��
     Minute As String                    ''��
 End Type


 '<���R�[�h�w�b�_>
 Private Type RECORD_ID
     RecordSpec As String                ''���R�[�h���
     DataKubun As String                 ''�f�[�^�敪
     MakeDate As YMD                     ''�f�[�^�쐬�N����
 End Type


 '<�������ʏ��P>
 Private Type RACE_ID
     Year As String                      ''�J�ÔN
     MonthDay As String                  ''�J�Ì���
     JyoCD As String                     ''���n��R�[�h
     Kaiji As String                     ''�J�É�[��N��]
     Nichiji As String                   ''�J�Ó���[N����]
     racenum As String                   ''���[�X�ԍ�
 End Type


 '<�������ʏ��Q>
 Private Type RACE_ID2
     Year As String                      ''�J�ÔN
     MonthDay As String                  ''�J�Ì���
     JyoCD As String                     ''���n��R�[�h
     Kaiji As String                     ''�J�É�[��N��]
     Nichiji As String                   ''�J�Ó���[N����]
 End Type


 '<���񐔁i�T�C�Y3byte�j>
 Private Type CHAKUKAISU3_INFO
     Chakukaisu(5) As String
 End Type


 '<���񐔁i�T�C�Y6byte�j>
 Private Type CHAKUKAISU6_INFO
     Chakukaisu(5) As String
 End Type


 '<�{�N�E�݌v���я��>
 Private Type SEI_RUIKEI_INFO
     SetYear As String                   ''�ݒ�N
     HonSyokinTotal As String            ''�{�܋����v
     Fukasyokin As String                ''�t���܋����v
     Chakukaisu(5) As String             ''����
 End Type


 '<�ŋߏd�܏������>
 Private Type SAIKIN_JYUSYO_INFO
     SaikinJyusyoid As RACE_ID           ''<�N��������R>
     Hondai As String                    ''�������{��
     Ryakusyo10 As String                ''����������10��
     Ryakusyo6 As String                 ''����������6��
     Ryakusyo3 As String                 ''����������3��
     GradeCD As String                   ''�O���[�h�R�[�h
     SyussoTosu As String                ''�o������
     KettoNum As String                  ''�����o�^�ԍ�
     Bamei As String                     ''�n��
 End Type


 '<�{�N�E�O�N�E�݌v���я��>
 Private Type HON_ZEN_RUIKEISEI_INFO
     SetYear As String                          ''�ݒ�N
     HonSyokinHeichi As String                  ''���n�{�܋����v
     HonSyokinSyogai As String                  ''��Q�{�܋����v
     FukaSyokinHeichi As String                 ''���n�t���܋����v
     FukaSyokinSyogai As String                 ''��Q�t���܋����v
     ChakuKaisuHeichi As CHAKUKAISU6_INFO       ''���n����
     ChakuKaisuSyogai As CHAKUKAISU6_INFO       ''��Q����
     ChakuKaisuJyo(19) As CHAKUKAISU6_INFO      ''���n��ʒ���
     ChakuKaisuKyori(5) As CHAKUKAISU6_INFO     ''�����ʒ���
 End Type


 '<���[�X���>
 Private Type RACE_INFO
     YoubiCD As String                   ''�j���R�[�h
     TokuNum As String                   ''���ʋ����ԍ�
     Hondai As String                    ''�������{��
     Fukudai As String                   ''����������
     Kakko As String                     ''�������J�b�R��
     HondaiEng As String                 ''�������{�艢��
     FukudaiEng As String                ''���������艢��
     KakkoEng As String                  ''�������J�b�R������
     Ryakusyo10 As String                ''���������̂P�O��
     Ryakusyo6 As String                 ''���������̂U��
     Ryakusyo3 As String                 ''���������̂R��
     Kubun As String                     ''�������敪
     Nkai As String                      ''�d�܉�[��N��]
 End Type


 '<�V��E�n����>
 Private Type TENKO_BABA_INFO
     TenkoCD As String                   ''�V��R�[�h
     SibaBabaCD As String                ''�Ŕn���ԃR�[�h
     DirtBabaCD As String                ''�_�[�g�n���ԃR�[�h
 End Type


 '<��������>
 Private Type RACE_JYOKEN
     SyubetuCD As String                 ''������ʃR�[�h
     KigoCD As String                    ''�����L���R�[�h
     JyuryoCD As String                  ''�d�ʎ�ʃR�[�h
     JyokenCD(4) As String               ''���������R�[�h
 End Type

 '''''''''''''''''''' �f�[�^�\���� ''''''''''''''''''''

'****** �P�D���ʓo�^�n ****************************************
 
 '<�o�^�n�����>
 Private Type TOKUUMA_INFO
     Num As String                       ''�A��
     KettoNum As String                  ''�����o�^�ԍ�
     Bamei As String                     ''�n��
     UmaKigoCD As String                 ''�n�L���R�[�h
     SexCD As String                     ''���ʃR�[�h
     TozaiCD As String                   ''�����t���������R�[�h
     ChokyosiCode As String              ''�����t�R�[�h
     ChokyosiRyakusyo As String          ''�����t������
     Futan As String                     ''���S�d��
     Koryu As String                     ''�𗬋敪
 End Type

 Public Type JV_TK_TOKUUMA
     head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
     id As RACE_ID                       ''<�������ʏ��P>
     RaceInfo As RACE_INFO               ''<���[�X���>
     GradeCD As String                   ''�O���[�h�R�[�h
     JyokenInfo As RACE_JYOKEN           ''<���������R�[�h>
     Kyori As String                     ''����
     TrackCD As String                   ''�g���b�N�R�[�h
     CourseKubunCD As String             ''�R�[�X�敪
     HandiDate As YMD                    ''�n���f���\��
     TorokuTosu As String                ''�o�^����
     TokuUmaInfo(299) As TOKUUMA_INFO    ''<�o�^�n�����>
     crlf As String                      ''���R�[�h���
     
 End Type

 '****** �Q�D���[�X�ڍ� ****************************************

 '<�R�[�i�[�ʉߏ���>
 Private Type CORNER_INFO
     Corner As String                    ''�R�[�i�[
     Syukaisu As String                  ''����
     Jyuni As String                    ''�e�ʉߏ���
    
 End Type

 Public Type JV_RA_RACE
     head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
     id As RACE_ID                       ''<�������ʏ��P>
     RaceInfo As RACE_INFO               ''<���[�X���>
     GradeCD As String                   ''�O���[�h�R�[�h
     GradeCDBefore As String             ''�ύX�O�O���[�h�R�[�h
     JyokenInfo As RACE_JYOKEN           ''<���������R�[�h>
     JyokenName As String                ''������������
     Kyori As String                     ''����
     KyoriBefore As String               ''�ύX�O����
     TrackCD As String                    ''�g���b�N�R�[�h
     TrackCDBefore As String             ''�ύX�O�g���b�N�R�[�h
     CourseKubunCD As String             ''�R�[�X�敪
     CourseKubunCDBefore As String       ''�ύX�O�R�[�X�敪
     Honsyokin(6) As String              ''�{�܋�
     HonsyokinBefore(4) As String        ''�ύX�O�{�܋�
     Fukasyokin(4) As String             ''�t���܋�
     FukasyokinBefore(2) As String       ''�ύX�O�t���܋�
     HassoTime As String                 ''��������
     HassoTimeBefore As String           ''�ύX�O��������
     TorokuTosu As String                ''�o�^����
     SyussoTosu As String                ''�o������
     NyusenTosu As String                ''��������
     TenkoBaba As TENKO_BABA_INFO        ''�V��E�n���ԃR�[�h
     LapTime(24) As String               ''���b�v�^�C��
     SyogaiMileTime As String            ''��Q�}�C���^�C��
     HaronTimeS3 As String               ''�O�R�n�����^�C��
     HaronTimeS4 As String               ''�O�S�n�����^�C��
     HaronTimeL3 As String               ''��R�n�����^�C��
     HaronTimeL4 As String               ''��S�n�����^�C��
     CornerInfo(3) As CORNER_INFO        ''<�R�[�i�[�ʉߏ���>
     RecordUpKubun As String             ''���R�[�h�X�V�敪
     crlf As String                      ''���R�[�h��؂�
 End Type


 '****** �R�D�n�����[�X��� ****************************************

 '<1���n(����n)���>
 Private Type CHAKUUMA_INFO
     KettoNum As String                  ''�����o�^�ԍ�
     Bamei As String                     ''�n��
 End Type

 Public Type JV_SE_RACE_UMA
     head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
     id As RACE_ID                       ''<�������ʏ��P>
     Wakuban As String                   ''�g��
     Umaban As String                    ''�n��
     KettoNum As String                  ''�����o�^�ԍ�
     Bamei As String                     ''�n��
     UmaKigoCD As String                 ''�n�L���R�[�h
     SexCD As String                     ''���ʃR�[�h
     HinsyuCD As String                  ''�i��R�[�h
     KeiroCD As String                   ''�ѐF�R�[�h
     Barei As String                     ''�n��
     TozaiCD As String                   ''���������R�[�h
     ChokyosiCode As String              ''�����t�R�[�h
     ChokyosiRyakusyo As String          ''�����t������
     BanusiCode As String                ''�n��R�[�h
     BanusiName As String                ''�n�喼
     Fukusyoku As String                 ''���F�W��
     reserved1 As String                 ''�\��
     Futan As String                     ''���S�d��
     FutanBefore As String               ''�ύX�O���S�d��
     Blinker As String                   ''�u�����J�[�g�p�敪
     reserved2 As String                 ''�\��
     KisyuCode As String                 ''�R��R�[�h
     KisyuCodeBefore As String           ''�ύX�O�R��R�[�h
     KisyuRyakusyo As String             ''�R�薼����
     KisyuRyakusyoBefore As String       ''�ύX�O�R�薼����
     MinaraiCD As String                 ''�R�茩�K�R�[�h
     MinaraiCDBefore As String           ''�ύX�O�R�茩�K�R�[�h
     BaTaijyu As String                  ''�n�̏d
     ZogenFugo As String                 ''��������
     ZogenSa As String                   ''������
     IJyoCD As String                    ''�ُ�敪�R�[�h
     NyusenJyuni As String               ''��������
     KakuteiJyuni As String              ''�m�蒅��
     DochakuKubun As String              ''�����敪
     DochakuTosu As String               ''��������
     Time As String                      ''���j�^�C��
     ChakusaCD As String                 ''�����R�[�h
     ChakusaCDP As String                ''+�����R�[�h
     ChakusaCDPP As String               ''++�����R�[�h
     Jyuni1c As String                   ''1�R�[�i�[�ł̏���
     Jyuni2c As String                   ''2�R�[�i�[�ł̏���
     Jyuni3c As String                   ''3�R�[�i�[�ł̏���
     Jyuni4c As String                   ''4�R�[�i�[�ł̏���
     Odds As String                      ''�P���I�b�Y
     Ninki As String                     ''�P���l�C��
     Honsyokin As String                 ''�l���{�܋�
     Fukasyokin As String                ''�l���t���܋�
     reserved3 As String                 ''�\��
     reserved4 As String                 ''�\��
     HaronTimeL4 As String               ''��S�n�����^�C��
     HaronTimeL3 As String               ''��R�n�����^�C��
     ChakuUmaInfo(2) As CHAKUUMA_INFO    ''<1���n(����n)���>
     TimeDiff As String                  ''�^�C����
     RecordUpKubun As String             ''���R�[�h�X�V�敪
     DMKubun As String                   ''�}�C�j���O�敪
     DMTime As String                    ''�}�C�j���O�\�z���j�^�C��
     DMGosaP As String                   ''�\���덷(�M���x)�{
     DMGosaM As String                   ''�\���덷(�M���x)�|
     DMJyuni As String                   ''�}�C�j���O�\�z����
     KyakusituKubun As String            ''���񃌁[�X�r������
     crlf As String                      ''���R�[�h��؂�
 End Type


 '****** �S�D���� ****************************************

 '<���ߏ��P �P�E���E�g>
 Private Type PAY_INFO1
     Umaban As String                    ''�n��
     Pay As String                       ''���ߋ�
     Ninki As String                     ''�l�C��
 End Type

 '<���ߏ��Q �n�A�E���C�h�E�\���E�n�P>
 Private Type PAY_INFO2
     Kumi As String                      ''�g��
     Pay As String                       ''���ߋ�
     Ninki As String                     ''�l�C��
 End Type

 '<���ߏ��R �R�A��>
 Private Type PAY_INFO3
     Kumi As String                      ''�g��
     Pay As String                       ''���ߋ�
     Ninki As String                     ''�l�C��
 End Type

 '<���ߏ��S �\��>
 Private Type PAY_INFO4
     Kumi As String                      ''�g��
     Pay As String                       ''���ߋ�
     Ninki As String                     ''�l�C��
 End Type

 Public Type JV_HR_PAY
     head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
     id As RACE_ID                       ''<�������ʏ��P>
     TorokuTosu As String                ''�o�^����
     SyussoTosu As String                ''�o������
     FuseirituFlag(8) As String          ''�s�����t���O
     TokubaraiFlag(8) As String          ''�����t���O
     HenkanFlag(8) As String             ''�Ԋ҃t���O
     HenkanUma(27) As String             ''�ԊҔn�ԏ��(�n��01�`28)
     HenkanWaku(7) As String             ''�ԊҘg�ԏ��(�g��1�`8)
     HenkanDoWaku(7) As String           ''�Ԋғ��g���(�g��1�`8)
     PayTansyo(2) As PAY_INFO1           ''<�P������>
     PayFukusyo(4) As PAY_INFO1          ''<��������>
     PayWakuren(2) As PAY_INFO1          ''<�g�A����>
     PayUmaren(2) As PAY_INFO2           ''<�n�A����>
     PayWide(6) As PAY_INFO2             ''<���C�h����>
     PayReserved1(2) As PAY_INFO2        ''<�\��>
     PayUmatan(5) As PAY_INFO2           ''<�n�P����>
     PaySanrenpuku(2) As PAY_INFO3       ''<3�A������>
     PaySanrentan(5) As PAY_INFO3        ''<3�A�P����>
     crlf As String                      ''���R�[�h��؂�
 End Type


 '****** �T�D�[���i�S�|���j****************************************

 '<�[�����P �P�E���E�g>
 Private Type HYO_INFO1
     Umaban As String                    ''�n��
     Hyo As String                       ''�[��
     Ninki As String                     ''�l�C
 End Type

 '<�[�����Q �n�A�E���C�h�E�n�P>
 Private Type HYO_INFO2
     Kumi As String                      ''�g��
     Hyo As String                       ''�[��
     Ninki As String                     ''�l�C
 End Type

 '<�[�����R �R�A���[��>
 Private Type HYO_INFO3
     Kumi As String                      ''�g��
     Hyo As String                       ''�[��
     Ninki As String                     ''�l�C
 End Type

 '<�[�����S �\��>
 Private Type HYO_INFO4
     Kumi As String                      ''�g��
     Hyo As String                       ''�[��
     Ninki As String                     ''�l�C
 End Type

 Public Type JV_H1_HYOSU_ZENKAKE
     head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
     id As RACE_ID                       ''<�������ʏ��P>
     TorokuTosu As String                ''�o�^����
     SyussoTosu As String                ''�o������
     HatubaiFlag(6) As String            ''�����t���O
     FukuChakuBaraiKey As String         ''���������L�[
     HenkanUma(27) As String             ''�ԊҔn�ԏ��(�n��01�`28)
     HenkanWaku(7) As String             ''�ԊҘg�ԏ��(�g��1�`8)
     HenkanDoWaku(7) As String           ''�Ԋғ��g���(�g��1�`8)
     HyoTansyo(27) As HYO_INFO1          ''<�P���[��>
     HyoFukusyo(27) As HYO_INFO1         ''<�����[��>
     HyoWakuren(35) As HYO_INFO1         ''<�g�A�[��>
     HyoUmaren(152) As HYO_INFO2         ''<�n�A�[��>
     HyoWide(152) As HYO_INFO2           ''<���C�h�[��>
     HyoUmatan(305) As HYO_INFO2         ''<�n�P�[��>
     HyoSanrenpuku(815) As HYO_INFO3     ''<3�A���[��>
     HyoTotal(13) As String              ''�[�����v
     crlf As String                      ''���R�[�h��؂�
 End Type


 Public Type JV_H6_HYOSU_SANRENTAN
     head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
     id As RACE_ID                       ''<�������ʏ��P>
     TorokuTosu As String                ''�o�^����
     SyussoTosu As String                ''�o������
     HatubaiFlag As String               ''�����t���O
     HenkanUma(17) As String             ''�ԊҔn�ԏ��(�n��01�`18)
     HyoSanrentan(4895) As HYO_INFO3     ''<3�A�P�[��>
     HyoTotal(2) As String               ''�[�����v
     crlf As String                      ''���R�[�h��؂�
 End Type

 '****** �U�D�I�b�Y�i�P���g�j****************************************

 '<�P���I�b�Y>
 Private Type ODDS_TANSYO_INFO
     Umaban As String                    ''�n��
     Odds As String                      ''�I�b�Y
     Ninki As String                     ''�l�C��
 End Type

 '<�����I�b�Y>
 Private Type ODDS_FUKUSYO_INFO
     Umaban As String                    ''�n��
     OddsLow As String                   ''�Œ�I�b�Y
     OddsHigh As String                  ''�ō��I�b�Y
     Ninki As String                     ''�l�C��
 End Type

 '<�g�A�I�b�Y>
 Private Type ODDS_WAKUREN_INFO
     Kumi As String                      ''�g
     Odds As String                      ''�I�b�Y
     Ninki As String                     ''�l�C��
 End Type

 Public Type JV_O1_ODDS_TANFUKUWAKU
     head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
     id As RACE_ID                       ''<�������ʏ��P>
     HappyoTime As MDHM                  ''���\��������
     TorokuTosu As String                ''�o�^����
     SyussoTosu As String                ''�o������
     TansyoFlag As String                ''�����t���O �P��
     FukusyoFlag As String               ''�����t���O ����
     WakurenFlag As String               ''�����t���O�@�g�A
     FukuChakuBaraiKey As String         ''���������L�[
     OddsTansyoInfo(27) As ODDS_TANSYO_INFO    ''<�P���I�b�Y>
     OddsFukusyoInfo(27) As ODDS_FUKUSYO_INFO  ''<�����[���I�b�Y>
     OddsWakurenInfo(35) As ODDS_WAKUREN_INFO  ''<�g�A�[���I�b�Y>
     TotalHyosuTansyo As String                ''�P���[�����v
     TotalHyosuFukusyo As String         ''�����[�����v
     TotalHyosuWakuren As String         ''�g�A�[�����v
     crlf As String                      ''���R�[�h��؂�
 End Type


 '****** �V�D�I�b�Y�i�n�A�j****************************************

 '<�n�A�I�b�Y>
 Private Type ODDS_UMAREN_INFO
     Kumi As String                      ''�g��
     Odds As String                      ''�I�b�Y
     Ninki As String                     ''�l�C��
 End Type

 Public Type JV_O2_ODDS_UMAREN
     head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
     id As RACE_ID                       ''<�������ʏ��P>
     HappyoTime As MDHM                  ''���\��������
     TorokuTosu As String                ''�o�^����
     SyussoTosu As String                ''�o������
     UmarenFlag As String                ''�����t���O�@�n�A
     OddsUmarenInfo(152) As ODDS_UMAREN_INFO   ''<�n�A�I�b�Y>
     TotalHyosuUmaren As String          ''�n�A�[�����v
     crlf As String                      ''���R�[�h��؂�
 End Type


 '****** �W�D�I�b�Y�i���C�h�j****************************************

 '<���C�h�I�b�Y>
 Private Type ODDS_WIDE_INFO
     Kumi As String                      ''�g��
     OddsLow As String                   ''�Œ�I�b�Y
     OddsHigh As String                  ''�ō��I�b�Y
     Ninki As String                     ''�l�C��
 End Type

 Public Type JV_O3_ODDS_WIDE
     head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
     id As RACE_ID                       ''<�������ʏ��P>
     HappyoTime As MDHM                  ''���\��������
     TorokuTosu As String                ''�o�^����
     SyussoTosu As String                ''�o������
     WideFlag As String                  ''�����t���O�@���C�h
     OddsWideInfo(152) As ODDS_WIDE_INFO ''<���C�h�I�b�Y>
     TotalHyosuWide As String            ''���C�h�[�����v
     crlf As String                      ''���R�[�h��؂�
 End Type


 '****** �X�D�I�b�Y�i�n�P�j ****************************************

 '<�n�P�I�b�Y>
 Private Type ODDS_UMATAN_INFO
     Kumi As String                      ''�g��
     Odds As String                      ''�I�b�Y
     Ninki As String                     ''�l�C��
 End Type

 Public Type JV_O4_ODDS_UMATAN
     head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
     id As RACE_ID                       ''<�������ʏ��P>
     HappyoTime As MDHM                  ''���\��������
     TorokuTosu As String                ''�o�^����
     SyussoTosu As String                ''�o������
     UmatanFlag As String                ''�����t���O�@�n�P
     OddsUmatanInfo(305) As ODDS_UMATAN_INFO ''<�n�P�I�b�Y>
     TotalHyosuUmatan As String          ''�n�P�[�����v
     crlf As String                      ''���R�[�h��؂�
 End Type


 '****** �P�O�D�I�b�Y�i�R�A���j***************************************

 '<3�A���I�b�Y>
 Private Type ODDS_SANREN_INFO
     Kumi As String                      ''�g��
     Odds As String                      ''�I�b�Y
     Ninki As String                     ''�l�C��
 End Type

 Public Type JV_O5_ODDS_SANREN
     head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
     id As RACE_ID                       ''<�������ʏ��P>
     HappyoTime As MDHM                  ''���\��������
     TorokuTosu As String                ''�o�^����
     SyussoTosu As String                ''�o������
     SanrenpukuFlag As String            ''�����t���O�@3�A��
     OddsSanrenInfo(815) As ODDS_SANREN_INFO ''<3�A���I�b�Y>
     TotalHyosuSanrenpuku As String          ''3�A���[�����v
     crlf As String                          ''���R�[�h��؂�
 End Type


 '****** �P�O�|�P�D�I�b�Y�i�R�A�P�j***************************************

 '<3�A�P�I�b�Y>
 Private Type ODDS_SANRENTAN_INFO
     Kumi As String                      ''�g��
     Odds As String                      ''�I�b�Y
     Ninki As String                     ''�l�C��
 End Type

 Public Type JV_O6_ODDS_SANRENTAN
     head As RECORD_ID                                                          ''<���R�[�h�w�b�_�[>
     id As RACE_ID                                                                      ''<�������ʏ��P>
     HappyoTime As MDHM                                                         ''���\��������
     TorokuTosu As String                                                       ''�o�^����
     SyussoTosu As String                                                       ''�o������
     SanrentanFlag As String                                    ''�����t���O�@3�A�P
     OddsSanrentanInfo(4895) As ODDS_SANRENTAN_INFO     ''<3�A�P�I�b�Y>
     TotalHyosuSanrentan As String                                      ''3�A�P�[�����v
     crlf As String                                                                     ''���R�[�h��؂�
 End Type
 
  Public Type JV_O6_ODDS_SANRENTAN2
     head As RECORD_ID                                                          ''<���R�[�h�w�b�_�[>
     id As RACE_ID                                                                      ''<�������ʏ��P>
     HappyoTime As MDHM                                                         ''���\��������
     TorokuTosu As String                                                       ''�o�^����
     SyussoTosu As String                                                       ''�o������
     SanrentanFlag As String                                    ''�����t���O�@3�A�P
     OddsSanrentanInfo As New Collection      ''<3�A�P�I�b�Y>
     TotalHyosuSanrentan As String                                      ''3�A�P�[�����v
     crlf As String                                                                     ''���R�[�h��؂�
 End Type


 '****** �P�P�D�����n�}�X�^ ****************************************

 '<�R�㌌�����>
 Private Type KETTO3_INFO
     HansyokuNum As String               ''�ɐB�o�^�ԍ�
     Bamei As String                     ''�n��
 End Type

 Public Type JV_UM_UMA
     head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
     KettoNum As String                  ''�����o�^�ԍ�
     DelKubun As String                  ''�����n�����敪
     RegDate As YMD                      ''�����n�o�^�N����
     DelDate As YMD                      ''�����n�����N����
     BirthDate As YMD                    ''���N����
     Bamei As String                     ''�n��
     BameiKana As String                 ''�n�����p�J�i
     BameiEng As String                  ''�n������
     UmaKigoCD As String                 ''�n�L���R�[�h
     SexCD As String                     ''���ʃR�[�h
     HinsyuCD As String                  ''�i��R�[�h
     KeiroCD As String                   ''�ѐF�R�[�h
     Ketto3Info(13) As KETTO3_INFO       ''<3�㌌�����>
     TozaiCD As String                   ''���������R�[�h
     ChokyosiCode As String              ''�����t�R�[�h
     ChokyosiRyakusyo As String          ''�����t������
     Syotai As String                    ''���Ғn�於
     BreederCode As String               ''���Y�҃R�[�h
     BreederName As String              ''���Y�Җ�
     SanchiName As String                ''�Y�n��
     BanusiCode As String                ''�n��R�[�h
     BanusiName As String                ''�n�喼
     RuikeiHonsyoHeiti As String         ''���n�{�܋��݌v
     RuikeiHonsyoSyogai As String        ''��Q�{�܋��݌v
     RuikeiFukaHeichi As String          ''���n�t���܋��݌v
     RuikeiFukaSyogai As String          ''��Q�t���܋��݌v
     RuikeiSyutokuHeichi As String       ''���n�����܋��݌v
     RuikeiSyutokuSyogai As String       ''��Q�����܋��݌v
     ChakuSogo As CHAKUKAISU3_INFO       ''��������
     ChakuChuo As CHAKUKAISU3_INFO       ''�������v����
     ChakuKaisuBa(6) As CHAKUKAISU3_INFO ''�n��ʒ���
     ChakuKaisuJyotai(11) As CHAKUKAISU3_INFO      ''�n���ԕʒ���
     ChakuKaisuKyori(5) As CHAKUKAISU3_INFO        ''�����ʒ���
     Kyakusitu(3) As String              ''�r���X��
     RaceCount As String                 ''�o�^���[�X��
     crlf As String                      ''���R�[�h��؂�
 End Type


 '****** �P�Q�D�R��}�X�^ ****************************************

 '<���R����>
 Private Type HATUKIJYO_INFO
     Hatukijyoid As RACE_ID              ''�N��������R
     SyussoTosu As String                ''�o������
     KettoNum As String                  ''�����o�^�ԍ�
     Bamei As String                     ''�n��
     KakuteiJyuni As String              ''�m�蒅��
     IJyoCD As String                    ''�ُ�敪�R�[�h
 End Type

 '<���������>
 Private Type HATUSYORI_INFO
     Hatusyoriid As RACE_ID              ''�N��������R
     SyussoTosu As String                ''�o������
     KettoNum As String                  ''�����o�^�ԍ�
     Bamei As String                     ''�n��
 End Type

 Public Type JV_KS_KISYU
     head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
     KisyuCode As String                 ''�R��R�[�h
     DelKubun As String                  ''�R�薕���敪
     IssueDate As YMD                    ''�R��Ƌ���t�N����
     DelDate As YMD                      ''�R��Ƌ������N����
     BirthDate As YMD                    ''���N����
     KisyuName As String                 ''�R�薼����
     reserved As String                  ''�\��
     KisyuNameKana As String             ''�R�薼���p�J�i
     KisyuRyakusyo As String             ''�R�薼����
     KisyuNameEng As String              ''�R�薼����
     SexCD As String                     ''���ʋ敪
     SikakuCD As String                  ''�R�掑�i�R�[�h
     MinaraiCD As String                 ''�R�茩�K�R�[�h
     TozaiCD As String                   ''�R�蓌�������R�[�h
     Syotai As String                    ''���Ғn�於
     ChokyosiCode As String              ''���������t�R�[�h
     ChokyosiRyakusyo As String          ''���������t������
     HatuKiJyo(1) As HATUKIJYO_INFO      ''<���R����>
     HatuSyori(1) As HATUSYORI_INFO      ''<���������>
     SaikinJyusyo(2) As SAIKIN_JYUSYO_INFO     ''<�ŋߏd�܏������>
     HonZenRuikei(2) As HON_ZEN_RUIKEISEI_INFO ''<�{�N�E�O�N�E�݌v���я��>
     crlf As String                           ''���R�[�h��؂�
 End Type


 '****** �P�R�D�����t�}�X�^ ****************************************

 Public Type JV_CH_CHOKYOSI
     head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
     ChokyosiCode As String              ''�����t�R�[�h
     DelKubun As String                  ''�����t�����敪
     IssueDate As YMD                    ''�����t�Ƌ���t�N����
     DelDate As YMD                      ''�����t�Ƌ������N����
     BirthDate As YMD                    ''���N����
     ChokyosiName As String              ''�����t������
     ChokyosiNameKana As String          ''�����t�����p�J�i
     ChokyosiRyakusyo As String          ''�����t������
     ChokyosiNameEng As String           ''�����t������
     SexCD As String                     ''���ʋ敪
     TozaiCD As String                   ''�����t���������R�[�h
     Syotai As String                    ''���Ғn�於
     SaikinJyusyo(2) As SAIKIN_JYUSYO_INFO     ''<�ŋߏd�܏������>
     HonZenRuikei(2) As HON_ZEN_RUIKEISEI_INFO ''<�{�N�E�O�N�E�݌v���я��>
     crlf As String                      ''���R�[�h��؂�
 End Type


 '******�P�S�D���Y�҃}�X�^ ****************************************

 Public Type JV_BR_BREEDER
     head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
     BreederCode As String               ''���Y�҃R�[�h
     BreederName_Co As String            ''���Y�Җ��i�@�l�i�L�j
     BreederName As String               ''���Y�Җ��i�@�l�i���j
     BreederNameKana As String           ''���Y�Җ����p�J�i
     BreederNameEng As String            ''���Y�Җ�����
     Address As String                   ''���Y�ҏZ�������Ȗ�
     HonRuikei(1) As SEI_RUIKEI_INFO     ''<�{�N�E�݌v���я��>
     crlf As String                      ''���R�[�h��؂�
 End Type


 '****** �P�T�D�n��}�X�^ ****************************************

 Public Type JV_BN_BANUSI
     head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
     BanusiCode As String                ''�n��R�[�h
     BanusiName_Co As String             ''�n�喼�i�@�l�i�L�j
     BanusiName As String                ''�n�喼�i�@�l�i���j
     BanusiNameKana As String            ''�n�喼���p�J�i
     BanusiNameEng As String             ''�n�喼����
     Fukusyoku As String                 ''���F�W��
     HonRuikei(1) As SEI_RUIKEI_INFO     ''<�{�N�E�݌v���я��>
     crlf As String                      ''���R�[�h��؂�
 End Type


 '****** �P�U�D�ɐB�n�}�X�^ ****************************************

 Public Type JV_HN_HANSYOKU
     head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
     HansyokuNum As String               ''�ɐB�o�^�ԍ�
     reserved As String                  ''�\��
     KettoNum As String                  ''�����o�^�ԍ�
     DelKubun As String                  ''�ɐB�n�����敪
     Bamei As String                     ''�n��
     BameiKana As String                 ''�n�����p�J�i
     BameiEng As String                  ''�n������
     BirthYear As String                 ''���N
     SexCD As String                     ''���ʃR�[�h
     HinsyuCD As String                  ''�i��R�[�h
     KeiroCD As String                   ''�ѐF�R�[�h
     HansyokuMochiKubun As String        ''�ɐB�n�����敪
     ImportYear As String                ''�A���N
     SanchiName As String                ''�Y�n��
     HansyokuFNum As String              ''���n�ɐB�o�^�ԍ�
     HansyokuMNum As String              ''��n�ɐB�o�^�ԍ�
     crlf As String                      ''���R�[�h��؂�
 End Type


 '****** �P�V�D�Y��}�X�^ ****************************************

 Public Type JV_SK_SANKU
     head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
     KettoNum As String                  ''�����o�^�ԍ�
     BirthDate As YMD                    ''���N����
     SexCD As String                     ''���ʃR�[�h
     HinsyuCD As String                  ''�i��R�[�h
     KeiroCD As String                   ''�ѐF�R�[�h
     SankuMochiKubun As String           ''�Y����敪
     ImportYear As String                ''�A���N
     BreederCode As String               ''���Y�҃R�[�h
     SanchiName As String                ''�Y�n��
     HansyokuNum(13) As String           ''3�㌌�� �ɐB�o�^�ԍ�
     crlf As String                      ''���R�[�h��؂�
 End Type


 '****** �P�W�D���R�[�h�}�X�^ ****************************************

 '<���R�[�h�ێ��n���>
 Private Type RECUMA_INFO
     KettoNum As String                  ''�����o�^�ԍ�
     Bamei As String                     ''�n��
     UmaKigoCD As String                 ''�n�L���R�[�h
     SexCD As String                     ''���ʃR�[�h
     ChokyosiCode As String              ''�����t�R�[�h
     ChokyosiName As String              ''�����t��
     Futan As String                     ''���S�d��
     KisyuCode As String                 ''�R��R�[�h
     KisyuName As String                 ''�R�薼
 End Type

 Public Type JV_RC_RECORD
     head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
     RecInfoKubun As String              ''���R�[�h���ʋ敪
     id As RACE_ID                       ''<�������ʏ��P>
     TokuNum As String                   ''���ʋ����ԍ�
     Hondai As String                    ''�������{��
     GradeCD As String                   ''�O���[�h�R�[�h
     SyubetuCD As String                 ''������ʃR�[�h
     Kyori As String                     ''����
     TrackCD As String                   ''�g���b�N�R�[�h
     RecKubun As String                  ''���R�[�h�敪
     RecTime As String                   ''���R�[�h�^�C��
     TenkoBaba As TENKO_BABA_INFO        ''�V��E�n����
     RecUmaInfo(2) As RECUMA_INFO        ''<���R�[�h�ێ��n���>
     crlf As String                      ''���R�[�h��؂�
 End Type


 '****** �P�X�D��H���� ****************************************

 Public Type JV_HC_HANRO
     head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
     TresenKubun As String               ''�g���Z���敪
     ChokyoDate As YMD                   ''�����N����
     ChokyoTime As String                ''��������
     KettoNum As String                  ''�����o�^�ԍ�
     HaronTime4 As String                ''4�n�����^�C�����v(800M-0M)
     LapTime4 As String                  ''���b�v�^�C��(800M-600M)
     HaronTime3 As String                ''3�n�����^�C�����v(600M-0M)
     LapTime3 As String                  ''���b�v�^�C��(600M-400M)
     HaronTime2 As String                ''2�n�����^�C�����v(400M-0M)
     LapTime2 As String                  ''���b�v�^�C��(400M-200M)
     LapTime1 As String                  ''���b�v�^�C��(200M-0M)
     crlf As String                      ''���R�[�h��؂�
 End Type


 '****** �Q�O�D�n�̏d ****************************************

 '<�n�̏d���>
 Private Type BATAIJYU_INFO
     Umaban As String                    ''�n��
     Bamei As String                     ''�n��
     BaTaijyu As String                  ''�n�̏d
     ZogenFugo As String                 ''��������
     ZogenSa As String                   ''������
 End Type

 Public Type JV_WH_BATAIJYU
     head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
     id As RACE_ID                       ''<�������ʏ��P>
     HappyoTime As MDHM                  ''���\��������
     BataijyuInfo(17) As BATAIJYU_INFO   ''<�n�̏d���>
     crlf As String                      ''���R�[�h��؂�
 End Type


 '****** �Q�P�D�V��n���� ******************************************

 Public Type JV_WE_WEATHER
     head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
     id As RACE_ID2                      ''<�������ʏ��Q>
     HappyoTime As MDHM                  ''���\��������
     HenkoID As String                   ''�ύX����
     TenkoBaba As TENKO_BABA_INFO        ''���ݏ�ԏ��
     TenkoBabaBefore As TENKO_BABA_INFO  ''�ύX�O��ԏ��
     crlf As String                      ''���R�[�h��؂�
    
 End Type

 '****** �Q�Q�D�o������E�������O ****************************************

 Public Type JV_AV_INFO
     head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
     id As RACE_ID                       ''<�������ʏ��P>
     HappyoTime As MDHM                  ''���\��������
     Umaban As String                    ''�n��
     Bamei As String                     ''�n��
     JiyuKubun As String                 ''���R�敪
     crlf As String                      ''���R�[�h��؂�
   
 End Type

 '************ �Q�R�D�R��ύX ****************************************

 '<�ύX���>
 Private Type JC_INFO
     Futan As String                     ''���S�d��
     KisyuCode As String                 ''�R��R�[�h
     KisyuName As String                 ''�R�薼
     MinaraiCD As String                 ''�R�茩�K�R�[�h
    
 End Type

 Public Type JV_JC_INFO
     head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
     id As RACE_ID                       ''<�������ʏ��P>
     HappyoTime As MDHM                  ''���\��������
     Umaban As String                    ''�n��
     Bamei As String                     ''�n��
     JCInfoAfter As JC_INFO              ''<�ύX����>
     JCInfoBefore As JC_INFO             ''<�ύX�O���>
     crlf As String                      ''���R�[�h��؂�
 End Type


 '************ �Q�R�|�P�D���������ύX ****************************************

 '<�ύX���>
 Private Type TC_INFO
     Ji As String                                               ''��
     Fun As String                                              ''��
 End Type

 Public Type JV_TC_INFO
     head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
     id As RACE_ID                       ''<�������ʏ��P>
     HappyoTime As MDHM                  ''���\��������
     TCInfoAfter As TC_INFO              ''<�ύX����>
     TCInfoBefore As TC_INFO             ''<�ύX�O���>
     crlf As String                      ''���R�[�h��؂�
 End Type


 '************ �Q�R�|�Q�D�R�[�X�ύX ****************************************

 '<�ύX���>
 Private Type CC_INFO
     Kyori As String                                    ''����
     TruckCD As String                                  ''�g���b�N�R�[�h
 End Type

 Public Type JV_CC_INFO
     head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
     id As RACE_ID                       ''<�������ʏ��P>
     HappyoTime As MDHM                  ''���\��������
     CCInfoAfter As CC_INFO              ''<�ύX����>
     CCInfoBefore As CC_INFO             ''<�ύX�O���>
     JiyuCD As String                    ''���R�R�[�h
     crlf As String                      ''���R�[�h��؂�
 End Type


 '****** �Q�S�D�f�[�^�}�C�j���O�\�z***********************************

 '<�}�C�j���O�\�z>
 Private Type DM_INFO
     Umaban As String                    ''�n��
     DMTime As String                    ''�\�z���j�^�C��
     DMGosaP As String                   ''�\�z�덷(�M���x)�{
     DMGosaM As String                   ''�\�z�덷(�M���x)�|
 End Type

 Public Type JV_DM_INFO
     head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
     id As RACE_ID                       ''<�������ʏ��P>
     MakeHM As HM                        ''�f�[�^�쐬����
     DMInfo(17) As DM_INFO               ''<�}�C�j���O�\�z>
     crlf As String                      ''���R�[�h��؂�
 End Type


 '****** �Q�T�D�J�ÃX�P�W���[��************************************

 '<�d�܈ē�>
 Private Type JYUSYO_INFO
     TokuNum As String                   ''���ʋ����ԍ�
     Hondai As String                    ''�������{��
     Ryakusyo10 As String                ''����������10��
     Ryakusyo6 As String                 ''����������6��
     Ryakusyo3 As String                 ''����������3��
     Nkai As String                      ''�d�܉�[��N��]
     GradeCD As String                   ''�O���[�h�R�[�h
     SyubetuCD As String                 ''������ʃR�[�h
     KigoCD As String                    ''�����L���R�[�h
     JyuryoCD As String                  ''�d�ʎ�ʃR�[�h
     Kyori As String                     ''����
     TrackCD As String                   ''�g���b�N�R�[�h
 End Type

 Public Type JV_YS_SCHEDULE
     head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
     id As RACE_ID2                      ''<�������ʏ��Q>
     YoubiCD As String                   ''�j���R�[�h
     JyusyoInfo(2) As JYUSYO_INFO        ''<�d�܈ē�>
     crlf As String                      ''���R�[�h��؂�
 End Type
 
     '''''''''''''''''''' �f�[�^�Z�b�g�֐� '''''''''''''''''''''''''''
    
   '****** �P�D���ʓo�^�n ****************************************
    
    Public Sub SetData_TK(ByRef lBuf As String, ByRef mBuf As JV_TK_TOKUUMA)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)              '' ���R�[�h���
            .DataKubun = IncMid(bytBuf, p, 1)               '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)                '' �N
                .Month = IncMid(bytBuf, p, 2)               '' ��
                .Day = IncMid(bytBuf, p, 2)                 '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)                    '' �J�ÔN
            .MonthDay = IncMid(bytBuf, p, 4)                '' �J�Ì���
            .JyoCD = IncMid(bytBuf, p, 2)                   '' ���n��R�[�h
            .Kaiji = IncMid(bytBuf, p, 2)                   '' �J�É�[��N��]
            .Nichiji = IncMid(bytBuf, p, 2)                 '' �J�Ó���[N����]
            .racenum = IncMid(bytBuf, p, 2)                 '' ���[�X�ԍ�
        End With ' id
        With .RaceInfo
            .YoubiCD = IncMid(bytBuf, p, 1)                 '' �j���R�[�h
            .TokuNum = IncMid(bytBuf, p, 4)                 '' ���ʋ����ԍ�
            .Hondai = IncMid(bytBuf, p, 60)                 '' �������{��
            .Fukudai = IncMid(bytBuf, p, 60)                '' ����������
            .Kakko = IncMid(bytBuf, p, 60)                  '' �������J�b�R��
            .HondaiEng = IncMid(bytBuf, p, 120)             '' �������{�艢��
            .FukudaiEng = IncMid(bytBuf, p, 120)            '' ���������艢��
            .KakkoEng = IncMid(bytBuf, p, 120)              '' �������J�b�R������
            .Ryakusyo10 = IncMid(bytBuf, p, 20)             '' ���������̂P�O��
            .Ryakusyo6 = IncMid(bytBuf, p, 12)              '' ���������̂U��
            .Ryakusyo3 = IncMid(bytBuf, p, 6)               '' ���������̂R��
            .Kubun = IncMid(bytBuf, p, 1)                   '' �������敪
            .Nkai = IncMid(bytBuf, p, 3)                    '' �d�܉�[��N��]
        End With ' RaceInfo
        .GradeCD = IncMid(bytBuf, p, 1)                     '' �O���[�h�R�[�h
        With .JyokenInfo
            .SyubetuCD = IncMid(bytBuf, p, 2)               '' ������ʃR�[�h
            .KigoCD = IncMid(bytBuf, p, 3)                  '' �����L���R�[�h
            .JyuryoCD = IncMid(bytBuf, p, 1)                '' �d�ʎ�ʃR�[�h
            For j = 0 To 4
                .JyokenCD(j) = IncMid(bytBuf, p, 3)         '' ���������R�[�h
            Next j
        End With ' JyokenInfo
        .Kyori = IncMid(bytBuf, p, 4)                       '' ����
        .TrackCD = IncMid(bytBuf, p, 2)                     '' �g���b�N�R�[�h
        .CourseKubunCD = IncMid(bytBuf, p, 2)               '' �R�[�X�敪
        With .HandiDate
            .Year = IncMid(bytBuf, p, 4)                    '' �N
            .Month = IncMid(bytBuf, p, 2)                   '' ��
            .Day = IncMid(bytBuf, p, 2)                     '' ��
        End With ' HandiDate
        .TorokuTosu = IncMid(bytBuf, p, 3)                  '' �o�^����
        For i = 0 To 299
            With .TokuUmaInfo(i)
                .Num = IncMid(bytBuf, p, 3)                 '' �A��
                .KettoNum = IncMid(bytBuf, p, 10)           '' �����o�^�ԍ�
                .Bamei = IncMid(bytBuf, p, 36)              '' �n��
                .UmaKigoCD = IncMid(bytBuf, p, 2)           '' �n�L���R�[�h
                .SexCD = IncMid(bytBuf, p, 1)               '' ���ʃR�[�h
                .TozaiCD = IncMid(bytBuf, p, 1)             '' �����t���������R�[�h
                .ChokyosiCode = IncMid(bytBuf, p, 5)        '' �����t�R�[�h
                .ChokyosiRyakusyo = IncMid(bytBuf, p, 8)    '' �����t������
                .Futan = IncMid(bytBuf, p, 3)               '' ���S�d��
                .Koryu = IncMid(bytBuf, p, 1)               '' �𗬋敪
            End With ' TokuUmaInfo
        Next i
        .crlf = IncMid(bytBuf, p, 2)                        '' ���R�[�h���
    End With

    '�o�b�t�@�̈���
    Erase bytBuf
       
    End Sub

    '****** �Q�D���[�X�ڍ� ****************************************
    Public Sub SetData_RA(ByRef lBuf As String, ByRef mBuf As JV_RA_RACE)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)              '' ���R�[�h���
            .DataKubun = IncMid(bytBuf, p, 1)               '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)                '' �N
                .Month = IncMid(bytBuf, p, 2)               '' ��
                .Day = IncMid(bytBuf, p, 2)                 '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)                    '' �J�ÔN
            .MonthDay = IncMid(bytBuf, p, 4)                '' �J�Ì���
            .JyoCD = IncMid(bytBuf, p, 2)                   '' ���n��R�[�h
            .Kaiji = IncMid(bytBuf, p, 2)                   '' �J�É�[��N��]
            .Nichiji = IncMid(bytBuf, p, 2)                 '' �J�Ó���[N����]
            .racenum = IncMid(bytBuf, p, 2)                 '' ���[�X�ԍ�
        End With ' id
        With .RaceInfo
            .YoubiCD = IncMid(bytBuf, p, 1)                 '' �j���R�[�h
            .TokuNum = IncMid(bytBuf, p, 4)                 '' ���ʋ����ԍ�
            .Hondai = IncMid(bytBuf, p, 60)                 '' �������{��
            .Fukudai = IncMid(bytBuf, p, 60)                '' ����������
            .Kakko = IncMid(bytBuf, p, 60)                  '' �������J�b�R��
            .HondaiEng = IncMid(bytBuf, p, 120)             '' �������{�艢��
            .FukudaiEng = IncMid(bytBuf, p, 120)            '' ���������艢��
            .KakkoEng = IncMid(bytBuf, p, 120)              '' �������J�b�R������
            .Ryakusyo10 = IncMid(bytBuf, p, 20)             '' ���������̂P�O��
            .Ryakusyo6 = IncMid(bytBuf, p, 12)              '' ���������̂U��
            .Ryakusyo3 = IncMid(bytBuf, p, 6)               '' ���������̂R��
            .Kubun = IncMid(bytBuf, p, 1)                   '' �������敪
            .Nkai = IncMid(bytBuf, p, 3)                    '' �d�܉�[��N��]
        End With ' RaceInfo
        .GradeCD = IncMid(bytBuf, p, 1)                     '' �O���[�h�R�[�h
        .GradeCDBefore = IncMid(bytBuf, p, 1)               '' �ύX�O�O���[�h�R�[�h
        With .JyokenInfo
            .SyubetuCD = IncMid(bytBuf, p, 2)               '' ������ʃR�[�h
            .KigoCD = IncMid(bytBuf, p, 3)                  '' �����L���R�[�h
            .JyuryoCD = IncMid(bytBuf, p, 1)                '' �d�ʎ�ʃR�[�h
            For j = 0 To 4
                .JyokenCD(j) = IncMid(bytBuf, p, 3)         '' ���������R�[�h
            Next j
        End With ' JyokenInfo
        .JyokenName = IncMid(bytBuf, p, 60)                 '' ������������
        .Kyori = IncMid(bytBuf, p, 4)                       '' ����
        .KyoriBefore = IncMid(bytBuf, p, 4)                 '' �ύX�O����
        .TrackCD = IncMid(bytBuf, p, 2)                     '' �g���b�N�R�[�h
        .TrackCDBefore = IncMid(bytBuf, p, 2)               '' �ύX�O�g���b�N�R�[�h
        .CourseKubunCD = IncMid(bytBuf, p, 2)               '' �R�[�X�敪
        .CourseKubunCDBefore = IncMid(bytBuf, p, 2)         '' �ύX�O�R�[�X�敪
        For i = 0 To 6
            .Honsyokin(i) = IncMid(bytBuf, p, 8)            '' �{�܋�
        Next i
        For i = 0 To 4
            .HonsyokinBefore(i) = IncMid(bytBuf, p, 8)      '' �ύX�O�{�܋�
        Next i
        For i = 0 To 4
            .Fukasyokin(i) = IncMid(bytBuf, p, 8)           '' �t���܋�
        Next i
        For i = 0 To 2
            .FukasyokinBefore(i) = IncMid(bytBuf, p, 8)     '' �ύX�O�t���܋�
        Next i
        .HassoTime = IncMid(bytBuf, p, 4)                   '' ��������
        .HassoTimeBefore = IncMid(bytBuf, p, 4)             '' �ύX�O��������
        .TorokuTosu = IncMid(bytBuf, p, 2)                  '' �o�^����
        .SyussoTosu = IncMid(bytBuf, p, 2)                  '' �o������
        .NyusenTosu = IncMid(bytBuf, p, 2)                  '' ��������
        With .TenkoBaba
            .TenkoCD = IncMid(bytBuf, p, 1)                 '' �V��R�[�h
            .SibaBabaCD = IncMid(bytBuf, p, 1)              '' �Ŕn���ԃR�[�h
            .DirtBabaCD = IncMid(bytBuf, p, 1)              '' �_�[�g�n���ԃR�[�h
        End With ' TenkoBaba
        For i = 0 To 24
            .LapTime(i) = IncMid(bytBuf, p, 3)              '' ���b�v�^�C��
        Next i
        .SyogaiMileTime = IncMid(bytBuf, p, 4)              '' ��Q�}�C���^�C��
        .HaronTimeS3 = IncMid(bytBuf, p, 3)                 '' �O�R�n�����^�C��
        .HaronTimeS4 = IncMid(bytBuf, p, 3)                 '' �O�S�n�����^�C��
        .HaronTimeL3 = IncMid(bytBuf, p, 3)                 '' ��R�n�����^�C��
        .HaronTimeL4 = IncMid(bytBuf, p, 3)                 '' ��S�n�����^�C��
        For i = 0 To 3
            With .CornerInfo(i)
                .Corner = IncMid(bytBuf, p, 1)              '' �R�[�i�[
                .Syukaisu = IncMid(bytBuf, p, 1)            '' ����
                .Jyuni = IncMid(bytBuf, p, 70)              '' �e�ʉߏ���
            End With ' CornerInfo
        Next i
        .RecordUpKubun = IncMid(bytBuf, p, 1)               '' ���R�[�h�X�V�敪
        .crlf = IncMid(bytBuf, p, 2)        '' ���R�[�h��؂�
    End With

    '�o�b�t�@�̈���
    Erase bytBuf
     
    End Sub


    '****** �R�D�n�����[�X��� ****************************************

    Public Sub SetData_SE(ByRef lBuf As String, ByRef mBuf As JV_SE_RACE_UMA)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMid(bytBuf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)        '' �N
                .Month = IncMid(bytBuf, p, 2)       '' ��
                .Day = IncMid(bytBuf, p, 2)         '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)            '' �J�ÔN
            .MonthDay = IncMid(bytBuf, p, 4)        '' �J�Ì���
            .JyoCD = IncMid(bytBuf, p, 2)           '' ���n��R�[�h
            .Kaiji = IncMid(bytBuf, p, 2)           '' �J�É�[��N��]
            .Nichiji = IncMid(bytBuf, p, 2)         '' �J�Ó���[N����]
            .racenum = IncMid(bytBuf, p, 2)         '' ���[�X�ԍ�
        End With ' id
        .Wakuban = IncMid(bytBuf, p, 1)             '' �g��
        .Umaban = IncMid(bytBuf, p, 2)              '' �n��
        .KettoNum = IncMid(bytBuf, p, 10)           '' �����o�^�ԍ�
        .Bamei = IncMid(bytBuf, p, 36)              '' �n��
        .UmaKigoCD = IncMid(bytBuf, p, 2)           '' �n�L���R�[�h
        .SexCD = IncMid(bytBuf, p, 1)               '' ���ʃR�[�h
        .HinsyuCD = IncMid(bytBuf, p, 1)            '' �i��R�[�h
        .KeiroCD = IncMid(bytBuf, p, 2)             '' �ѐF�R�[�h
        .Barei = IncMid(bytBuf, p, 2)               '' �n��
        .TozaiCD = IncMid(bytBuf, p, 1)             '' ���������R�[�h
        .ChokyosiCode = IncMid(bytBuf, p, 5)        '' �����t�R�[�h
        .ChokyosiRyakusyo = IncMid(bytBuf, p, 8)    '' �����t������
        .BanusiCode = IncMid(bytBuf, p, 6)          '' �n��R�[�h
        .BanusiName = IncMid(bytBuf, p, 64)         '' �n�喼
        .Fukusyoku = IncMid(bytBuf, p, 60)          '' ���F�W��
        .reserved1 = IncMid(bytBuf, p, 60)          '' �\��
        .Futan = IncMid(bytBuf, p, 3)               '' ���S�d��
        .FutanBefore = IncMid(bytBuf, p, 3)         '' �ύX�O���S�d��
        .Blinker = IncMid(bytBuf, p, 1)             '' �u�����J�[�g�p�敪
        .reserved2 = IncMid(bytBuf, p, 1)           '' �\��
        .KisyuCode = IncMid(bytBuf, p, 5)           '' �R��R�[�h
        .KisyuCodeBefore = IncMid(bytBuf, p, 5)     '' �ύX�O�R��R�[�h
        .KisyuRyakusyo = IncMid(bytBuf, p, 8)       '' �R�薼����
        .KisyuRyakusyoBefore = IncMid(bytBuf, p, 8) '' �ύX�O�R�薼����
        .MinaraiCD = IncMid(bytBuf, p, 1)           '' �R�茩�K�R�[�h
        .MinaraiCDBefore = IncMid(bytBuf, p, 1)     '' �ύX�O�R�茩�K�R�[�h
        .BaTaijyu = IncMid(bytBuf, p, 3)            '' �n�̏d
        .ZogenFugo = IncMid(bytBuf, p, 1)           '' ��������
        .ZogenSa = IncMid(bytBuf, p, 3)             '' ������
        .IJyoCD = IncMid(bytBuf, p, 1)              '' �ُ�敪�R�[�h
        .NyusenJyuni = IncMid(bytBuf, p, 2)         '' ��������
        .KakuteiJyuni = IncMid(bytBuf, p, 2)        '' �m�蒅��
        .DochakuKubun = IncMid(bytBuf, p, 1)        '' �����敪
        .DochakuTosu = IncMid(bytBuf, p, 1)         '' ��������
        .Time = IncMid(bytBuf, p, 4)                '' ���j�^�C��
        .ChakusaCD = IncMid(bytBuf, p, 3)           '' �����R�[�h
        .ChakusaCDP = IncMid(bytBuf, p, 3)          '' +�����R�[�h
        .ChakusaCDPP = IncMid(bytBuf, p, 3)         '' ++�����R�[�h
        .Jyuni1c = IncMid(bytBuf, p, 2)             '' 1�R�[�i�[�ł̏���
        .Jyuni2c = IncMid(bytBuf, p, 2)             '' 2�R�[�i�[�ł̏���
        .Jyuni3c = IncMid(bytBuf, p, 2)             '' 3�R�[�i�[�ł̏���
        .Jyuni4c = IncMid(bytBuf, p, 2)             '' 4�R�[�i�[�ł̏���
        .Odds = IncMid(bytBuf, p, 4)                '' �P���I�b�Y
        .Ninki = IncMid(bytBuf, p, 2)               '' �P���l�C��
        .Honsyokin = IncMid(bytBuf, p, 8)           '' �l���{�܋�
        .Fukasyokin = IncMid(bytBuf, p, 8)          '' �l���t���܋�
        .reserved3 = IncMid(bytBuf, p, 3)           '' �\��
        .reserved4 = IncMid(bytBuf, p, 3)           '' �\��
        .HaronTimeL4 = IncMid(bytBuf, p, 3)         '' ��S�n�����^�C��
        .HaronTimeL3 = IncMid(bytBuf, p, 3)         '' ��R�n�����^�C��
        For i = 0 To 2
            With .ChakuUmaInfo(i)
                .KettoNum = IncMid(bytBuf, p, 10)   '' �����o�^�ԍ�
                .Bamei = IncMid(bytBuf, p, 36)      '' �n��
            End With ' ChakuUmaInfo
        Next i
        .TimeDiff = IncMid(bytBuf, p, 4)            '' �^�C����
        .RecordUpKubun = IncMid(bytBuf, p, 1)       '' ���R�[�h�X�V�敪
        .DMKubun = IncMid(bytBuf, p, 1)             '' �}�C�j���O�敪
        .DMTime = IncMid(bytBuf, p, 5)              '' �}�C�j���O�\�z���j�^�C��
        .DMGosaP = IncMid(bytBuf, p, 4)             '' �\���덷(�M���x)�{
        .DMGosaM = IncMid(bytBuf, p, 4)             '' �\���덷(�M���x)�|
        .DMJyuni = IncMid(bytBuf, p, 2)             '' �}�C�j���O�\�z����
        .KyakusituKubun = IncMid(bytBuf, p, 1)      '' ���񃌁[�X�r������
        .crlf = IncMid(bytBuf, p, 2)                '' ���R�[�h��؂�
    End With

    '�o�b�t�@�̈���
    Erase bytBuf
   
    End Sub


    '****** �S�D���� ****************************************

    Public Sub SetData_HR(lBuf As String, ByRef mBuf As JV_HR_PAY)
    Dim bytBuf() As Byte                                    '' �o�C�g�z��ŏ������邽�߂̃o�b�t�@
    Dim i As Integer                                        '' ���[�v�J�E���^
    Dim j As Integer                                        '' ���[�v�J�E���^
    Dim k As Integer                                        '' ���[�v�J�E���^
    Dim p As Long                                           '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)              '' ���R�[�h���
            .DataKubun = IncMid(bytBuf, p, 1)               '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)                '' �N
                .Month = IncMid(bytBuf, p, 2)               '' ��
                .Day = IncMid(bytBuf, p, 2)                 '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)                    '' �J�ÔN
            .MonthDay = IncMid(bytBuf, p, 4)                '' �J�Ì���
            .JyoCD = IncMid(bytBuf, p, 2)                   '' ���n��R�[�h
            .Kaiji = IncMid(bytBuf, p, 2)                   '' �J�É�[��N��]
            .Nichiji = IncMid(bytBuf, p, 2)                 '' �J�Ó���[N����]
            .racenum = IncMid(bytBuf, p, 2)                 '' ���[�X�ԍ�
        End With ' id
        .TorokuTosu = IncMid(bytBuf, p, 2)                  '' �o�^����
        .SyussoTosu = IncMid(bytBuf, p, 2)                  '' �o������
        For i = 0 To 8
            .FuseirituFlag(i) = IncMid(bytBuf, p, 1)        '' �s�����t���O
        Next i
        For i = 0 To 8
            .TokubaraiFlag(i) = IncMid(bytBuf, p, 1)        '' �����t���O
        Next i
        For i = 0 To 8
            .HenkanFlag(i) = IncMid(bytBuf, p, 1)           '' �Ԋ҃t���O
        Next i
        For i = 0 To 27
            .HenkanUma(i) = IncMid(bytBuf, p, 1)            '' �ԊҔn�ԏ��(�n��01�`28)
        Next i
        For i = 0 To 7
            .HenkanWaku(i) = IncMid(bytBuf, p, 1)           '' �ԊҘg�ԏ��(�g��1�`8)
        Next i
        For i = 0 To 7
            .HenkanDoWaku(i) = IncMid(bytBuf, p, 1)         '' �Ԋғ��g���(�g��1�`8)
        Next i
        For i = 0 To 2
            With .PayTansyo(i)
                .Umaban = IncMid(bytBuf, p, 2)              '' �n��
                .Pay = IncMid(bytBuf, p, 9)                 '' ���ߋ�
                .Ninki = IncMid(bytBuf, p, 2)               '' �l�C��
            End With ' PayTansyo
        Next i
        For i = 0 To 4
            With .PayFukusyo(i)
                .Umaban = IncMid(bytBuf, p, 2)              '' �n��
                .Pay = IncMid(bytBuf, p, 9)                 '' ���ߋ�
                .Ninki = IncMid(bytBuf, p, 2)               '' �l�C��
            End With ' PayFukusyo
        Next i
        For i = 0 To 2
            With .PayWakuren(i)
                .Umaban = IncMid(bytBuf, p, 2)              '' �n��
                .Pay = IncMid(bytBuf, p, 9)                 '' ���ߋ�
                .Ninki = IncMid(bytBuf, p, 2)               '' �l�C��
            End With ' PayWakuren
        Next i
        For i = 0 To 2
            With .PayUmaren(i)
                .Kumi = IncMid(bytBuf, p, 4)                '' �g��
                .Pay = IncMid(bytBuf, p, 9)                 '' ���ߋ�
                .Ninki = IncMid(bytBuf, p, 3)               '' �l�C��
            End With ' PayUmaren
        Next i
        For i = 0 To 6
            With .PayWide(i)
                .Kumi = IncMid(bytBuf, p, 4)                '' �g��
                .Pay = IncMid(bytBuf, p, 9)                 '' ���ߋ�
                .Ninki = IncMid(bytBuf, p, 3)               '' �l�C��
            End With ' PayWide
        Next i
        For i = 0 To 2
            With .PayReserved1(i)
                .Kumi = IncMid(bytBuf, p, 4)                '' �g��
                .Pay = IncMid(bytBuf, p, 9)                 '' ���ߋ�
                .Ninki = IncMid(bytBuf, p, 3)               '' �l�C��
            End With ' PayReserved1
        Next i
        For i = 0 To 5
            With .PayUmatan(i)
                .Kumi = IncMid(bytBuf, p, 4)                '' �g��
                .Pay = IncMid(bytBuf, p, 9)                 '' ���ߋ�
                .Ninki = IncMid(bytBuf, p, 3)               '' �l�C��
            End With ' PayUmatan
        Next i
        For i = 0 To 2
            With .PaySanrenpuku(i)
                .Kumi = IncMid(bytBuf, p, 6)                '' �g��
                .Pay = IncMid(bytBuf, p, 9)                 '' ���ߋ�
                .Ninki = IncMid(bytBuf, p, 3)               '' �l�C��
            End With ' PaySanrenpuku
        Next i
        For i = 0 To 5
            With .PaySanrentan(i)
                .Kumi = IncMid(bytBuf, p, 6)                '' �g��
                .Pay = IncMid(bytBuf, p, 9)                 '' ���ߋ�
                .Ninki = IncMid(bytBuf, p, 4)               '' �l�C��
            End With ' PayReserved2
        Next i
        .crlf = IncMid(bytBuf, p, 2)                        '' ���R�[�h��؂�
    End With
    
    '�o�b�t�@�̈���
    Erase bytBuf
    
    End Sub

    '****** �T�D�[���i�S�|���j****************************************

    Public Sub SetData_H1(lBuf As String, ByRef mBuf As JV_H1_HYOSU_ZENKAKE)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMid(bytBuf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)        '' �N
                .Month = IncMid(bytBuf, p, 2)       '' ��
                .Day = IncMid(bytBuf, p, 2)         '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)            '' �J�ÔN
            .MonthDay = IncMid(bytBuf, p, 4)        '' �J�Ì���
            .JyoCD = IncMid(bytBuf, p, 2)           '' ���n��R�[�h
            .Kaiji = IncMid(bytBuf, p, 2)           '' �J�É�[��N��]
            .Nichiji = IncMid(bytBuf, p, 2)         '' �J�Ó���[N����]
            .racenum = IncMid(bytBuf, p, 2)         '' ���[�X�ԍ�
        End With ' id
        .TorokuTosu = IncMid(bytBuf, p, 2)          '' �o�^����
        .SyussoTosu = IncMid(bytBuf, p, 2)          '' �o������
        For i = 0 To 6
            .HatubaiFlag(i) = IncMid(bytBuf, p, 1)  '' �����t���O
        Next i
        .FukuChakuBaraiKey = IncMid(bytBuf, p, 1)   '' ���������L�[
        For i = 0 To 27
            .HenkanUma(i) = IncMid(bytBuf, p, 1)    '' �ԊҔn�ԏ��(�n��01�`28)
        Next i
        For i = 0 To 7
            .HenkanWaku(i) = IncMid(bytBuf, p, 1)   '' �ԊҘg�ԏ��(�g��1�`8)
        Next i
        For i = 0 To 7
            .HenkanDoWaku(i) = IncMid(bytBuf, p, 1) '' �Ԋғ��g���(�g��1�`8)
        Next i
        For i = 0 To 27
            With .HyoTansyo(i)
                .Umaban = IncMid(bytBuf, p, 2)      '' �n��
                .Hyo = IncMid(bytBuf, p, 11)        '' �[��
                .Ninki = IncMid(bytBuf, p, 2)       '' �l�C
            End With ' HyoTansyo
        Next i
        For i = 0 To 27
            With .HyoFukusyo(i)
                .Umaban = IncMid(bytBuf, p, 2)      '' �n��
                .Hyo = IncMid(bytBuf, p, 11)        '' �[��
                .Ninki = IncMid(bytBuf, p, 2)       '' �l�C
            End With ' HyoFukusyo
        Next i
        For i = 0 To 35
            With .HyoWakuren(i)
                .Umaban = IncMid(bytBuf, p, 2)      '' �n��
                .Hyo = IncMid(bytBuf, p, 11)        '' �[��
                .Ninki = IncMid(bytBuf, p, 2)       '' �l�C
            End With ' HyoWakuren
        Next i
        For i = 0 To 152
            With .HyoUmaren(i)
                .Kumi = IncMid(bytBuf, p, 4)        '' �g��
                .Hyo = IncMid(bytBuf, p, 11)        '' �[��
                .Ninki = IncMid(bytBuf, p, 3)       '' �l�C
            End With ' HyoUmaren
        Next i
        For i = 0 To 152
            With .HyoWide(i)
                .Kumi = IncMid(bytBuf, p, 4)        '' �g��
                .Hyo = IncMid(bytBuf, p, 11)        '' �[��
                .Ninki = IncMid(bytBuf, p, 3)       '' �l�C
            End With ' HyoWide
        Next i
        For i = 0 To 305
            With .HyoUmatan(i)
                .Kumi = IncMid(bytBuf, p, 4)        '' �g��
                .Hyo = IncMid(bytBuf, p, 11)        '' �[��
                .Ninki = IncMid(bytBuf, p, 3)       '' �l�C
            End With ' HyoUmatan
        Next i
        For i = 0 To 815
            With .HyoSanrenpuku(i)
                .Kumi = IncMid(bytBuf, p, 6)        '' �g��
                .Hyo = IncMid(bytBuf, p, 11)        '' �[��
                .Ninki = IncMid(bytBuf, p, 3)       '' �l�C
            End With ' HyoSanrenpuku
        Next i
        For i = 0 To 13
            .HyoTotal(i) = IncMid(bytBuf, p, 11)    '' �[�����v
        Next i
        .crlf = IncMid(bytBuf, p, 2)                '' ���R�[�h��؂�
    End With

    '�o�b�t�@�̈���
    Erase bytBuf
    
    End Sub


    '****** �U�D�I�b�Y�i�P���g�j****************************************

    Public Sub SetData_O1(lBuf As String, ByRef mBuf As JV_O1_ODDS_TANFUKUWAKU)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMid(bytBuf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)        '' �N
                .Month = IncMid(bytBuf, p, 2)       '' ��
                .Day = IncMid(bytBuf, p, 2) '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)            '' �J�ÔN
            .MonthDay = IncMid(bytBuf, p, 4)        '' �J�Ì���
            .JyoCD = IncMid(bytBuf, p, 2)           '' ���n��R�[�h
            .Kaiji = IncMid(bytBuf, p, 2)           '' �J�É�[��N��]
            .Nichiji = IncMid(bytBuf, p, 2)         '' �J�Ó���[N����]
            .racenum = IncMid(bytBuf, p, 2)         '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = IncMid(bytBuf, p, 2)           '' ��
            .Day = IncMid(bytBuf, p, 2)             '' ��
            .Hour = IncMid(bytBuf, p, 2)            '' ��
            .Minute = IncMid(bytBuf, p, 2)          '' ��
        End With ' HappyoTime
        .TorokuTosu = IncMid(bytBuf, p, 2)          '' �o�^����
        .SyussoTosu = IncMid(bytBuf, p, 2)          '' �o������
        .TansyoFlag = IncMid(bytBuf, p, 1)          '' �����t���O
        .FukusyoFlag = IncMid(bytBuf, p, 1)         '' �����t���O
        .WakurenFlag = IncMid(bytBuf, p, 1)         '' �����t���O�@�g�A
        .FukuChakuBaraiKey = IncMid(bytBuf, p, 1)   '' ���������L�[
        For i = 0 To 27
            With .OddsTansyoInfo(i)
                .Umaban = IncMid(bytBuf, p, 2)      '' �n��
                .Odds = IncMid(bytBuf, p, 4)        '' �I�b�Y
                .Ninki = IncMid(bytBuf, p, 2)       '' �l�C��
            End With ' OddsTansyoInfo
        Next i
        For i = 0 To 27
            With .OddsFukusyoInfo(i)
                .Umaban = IncMid(bytBuf, p, 2)      '' �n��
                .OddsLow = IncMid(bytBuf, p, 4)     '' �Œ�I�b�Y
                .OddsHigh = IncMid(bytBuf, p, 4)    '' �ō��I�b�Y
                .Ninki = IncMid(bytBuf, p, 2)       '' �l�C��
            End With ' OddsFukusyoInfo
        Next i
        For i = 0 To 35
            With .OddsWakurenInfo(i)
                .Kumi = IncMid(bytBuf, p, 2)        '' �g
                .Odds = IncMid(bytBuf, p, 5)        '' �I�b�Y
                .Ninki = IncMid(bytBuf, p, 2)       '' �l�C��
            End With ' OddsWakurenInfo
        Next i
        .TotalHyosuTansyo = IncMid(bytBuf, p, 11)   '' �P���[�����v
        .TotalHyosuFukusyo = IncMid(bytBuf, p, 11)  '' �����[�����v
        .TotalHyosuWakuren = IncMid(bytBuf, p, 11)  '' �g�A�[�����v
        .crlf = IncMid(bytBuf, p, 2)                '' ���R�[�h��؂�
    End With

    '�o�b�t�@�̈���
    Erase bytBuf

    End Sub


    '****** �V�D�I�b�Y�i�n�A�j****************************************

    Public Sub SetData_O2(lBuf As String, ByRef mBuf As JV_O2_ODDS_UMAREN)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMid(bytBuf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)        '' �N
                .Month = IncMid(bytBuf, p, 2)       '' ��
                .Day = IncMid(bytBuf, p, 2) '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)    '' �J�ÔN
            .MonthDay = IncMid(bytBuf, p, 4)        '' �J�Ì���
            .JyoCD = IncMid(bytBuf, p, 2)   '' ���n��R�[�h
            .Kaiji = IncMid(bytBuf, p, 2)   '' �J�É�[��N��]
            .Nichiji = IncMid(bytBuf, p, 2) '' �J�Ó���[N����]
            .racenum = IncMid(bytBuf, p, 2) '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = IncMid(bytBuf, p, 2)   '' ��
            .Day = IncMid(bytBuf, p, 2)     '' ��
            .Hour = IncMid(bytBuf, p, 2)    '' ��
            .Minute = IncMid(bytBuf, p, 2)  '' ��
        End With ' HappyoTime
        .TorokuTosu = IncMid(bytBuf, p, 2)  '' �o�^����
        .SyussoTosu = IncMid(bytBuf, p, 2)  '' �o������
        .UmarenFlag = IncMid(bytBuf, p, 1)  '' �����t���O�@�n�A
        For i = 0 To 152
            With .OddsUmarenInfo(i)
                .Kumi = IncMid(bytBuf, p, 4)        '' �g��
                .Odds = IncMid(bytBuf, p, 6)        '' �I�b�Y
                .Ninki = IncMid(bytBuf, p, 3)       '' �l�C��
            End With ' OddsUmarenInfo
        Next i
        .TotalHyosuUmaren = IncMid(bytBuf, p, 11)   '' �n�A�[�����v
        .crlf = IncMid(bytBuf, p, 2)        '' ���R�[�h��؂�
    End With

    '�o�b�t�@�̈���
    Erase bytBuf

    End Sub


    '****** �W�D�I�b�Y�i���C�h�j****************************************

    Public Sub SetData_O3(lBuf As String, ByRef mBuf As JV_O3_ODDS_WIDE)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMid(bytBuf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)        '' �N
                .Month = IncMid(bytBuf, p, 2)       '' ��
                .Day = IncMid(bytBuf, p, 2) '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)    '' �J�ÔN
            .MonthDay = IncMid(bytBuf, p, 4)        '' �J�Ì���
            .JyoCD = IncMid(bytBuf, p, 2)   '' ���n��R�[�h
            .Kaiji = IncMid(bytBuf, p, 2)   '' �J�É�[��N��]
            .Nichiji = IncMid(bytBuf, p, 2) '' �J�Ó���[N����]
            .racenum = IncMid(bytBuf, p, 2) '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = IncMid(bytBuf, p, 2)   '' ��
            .Day = IncMid(bytBuf, p, 2)     '' ��
            .Hour = IncMid(bytBuf, p, 2)    '' ��
            .Minute = IncMid(bytBuf, p, 2)  '' ��
        End With ' HappyoTime
        .TorokuTosu = IncMid(bytBuf, p, 2)  '' �o�^����
        .SyussoTosu = IncMid(bytBuf, p, 2)  '' �o������
        .WideFlag = IncMid(bytBuf, p, 1)    '' �����t���O�@���C�h
        For i = 0 To 152
            With .OddsWideInfo(i)
                .Kumi = IncMid(bytBuf, p, 4)        '' �g��
                .OddsLow = IncMid(bytBuf, p, 5)     '' �Œ�I�b�Y
                .OddsHigh = IncMid(bytBuf, p, 5)    '' �ō��I�b�Y
                .Ninki = IncMid(bytBuf, p, 3)       '' �l�C��
            End With ' OddsWideInfo
        Next i
        .TotalHyosuWide = IncMid(bytBuf, p, 11)     '' ���C�h�[�����v
        .crlf = IncMid(bytBuf, p, 2)        '' ���R�[�h��؂�
    End With

    '�o�b�t�@�̈���
    Erase bytBuf
   
    End Sub


    '****** �X�D�I�b�Y�i�n�P�j ****************************************

    Public Sub SetData_O4(lBuf As String, ByRef mBuf As JV_O4_ODDS_UMATAN)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMid(bytBuf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)        '' �N
                .Month = IncMid(bytBuf, p, 2)       '' ��
                .Day = IncMid(bytBuf, p, 2)         '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)            '' �J�ÔN
            .MonthDay = IncMid(bytBuf, p, 4)        '' �J�Ì���
            .JyoCD = IncMid(bytBuf, p, 2)           '' ���n��R�[�h
            .Kaiji = IncMid(bytBuf, p, 2)           '' �J�É�[��N��]
            .Nichiji = IncMid(bytBuf, p, 2)         '' �J�Ó���[N����]
            .racenum = IncMid(bytBuf, p, 2)         '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = IncMid(bytBuf, p, 2)           '' ��
            .Day = IncMid(bytBuf, p, 2)             '' ��
            .Hour = IncMid(bytBuf, p, 2)            '' ��
            .Minute = IncMid(bytBuf, p, 2)          '' ��
        End With ' HappyoTime
        .TorokuTosu = IncMid(bytBuf, p, 2)          '' �o�^����
        .SyussoTosu = IncMid(bytBuf, p, 2)          '' �o������
        .UmatanFlag = IncMid(bytBuf, p, 1)          '' �����t���O�@�n�P
        For i = 0 To 305
            With .OddsUmatanInfo(i)
                .Kumi = IncMid(bytBuf, p, 4)        '' �g��
                .Odds = IncMid(bytBuf, p, 6)        '' �I�b�Y
                .Ninki = IncMid(bytBuf, p, 3)       '' �l�C��
            End With ' OddsUmatanInfo
        Next i
        .TotalHyosuUmatan = IncMid(bytBuf, p, 11)   '' �n�P�[�����v
        .crlf = IncMid(bytBuf, p, 2)                '' ���R�[�h��؂�
    End With

    '�o�b�t�@�̈���
    Erase bytBuf

    End Sub


    '****** �P�O�D�I�b�Y�i�R�A���j***************************************

    Public Sub SetData_O5(lBuf As String, ByRef mBuf As JV_O5_ODDS_SANREN)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMid(bytBuf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)        '' �N
                .Month = IncMid(bytBuf, p, 2)       '' ��
                .Day = IncMid(bytBuf, p, 2)         '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)            '' �J�ÔN
            .MonthDay = IncMid(bytBuf, p, 4)        '' �J�Ì���
            .JyoCD = IncMid(bytBuf, p, 2)           '' ���n��R�[�h
            .Kaiji = IncMid(bytBuf, p, 2)               '' �J�É�[��N��]
            .Nichiji = IncMid(bytBuf, p, 2)         '' �J�Ó���[N����]
            .racenum = IncMid(bytBuf, p, 2)         '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = IncMid(bytBuf, p, 2)           '' ��
            .Day = IncMid(bytBuf, p, 2)             '' ��
            .Hour = IncMid(bytBuf, p, 2)            '' ��
            .Minute = IncMid(bytBuf, p, 2)          '' ��
        End With ' HappyoTime
        .TorokuTosu = IncMid(bytBuf, p, 2)          '' �o�^����
        .SyussoTosu = IncMid(bytBuf, p, 2)          '' �o������
        .SanrenpukuFlag = IncMid(bytBuf, p, 1)      '' �����t���O�@3�A��
        For i = 0 To 815
            With .OddsSanrenInfo(i)
                .Kumi = IncMid(bytBuf, p, 6)        '' �g��
                .Odds = IncMid(bytBuf, p, 6)        '' �I�b�Y
                .Ninki = IncMid(bytBuf, p, 3)       '' �l�C��
            End With ' OddsSanrenInfo
        Next i
        .TotalHyosuSanrenpuku = IncMid(bytBuf, p, 11)       '' 3�A���[�����v
        .crlf = IncMid(bytBuf, p, 2)        '' ���R�[�h��؂�
    End With

    '�o�b�t�@�̈���
    Erase bytBuf
   
    End Sub


    '****** �P�P�D�����n�}�X�^ ****************************************

    Public Sub SetData_UM(ByVal lBuf As String, ByRef mBuf As JV_UM_UMA)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMid(bytBuf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)        '' �N
                .Month = IncMid(bytBuf, p, 2)       '' ��
                .Day = IncMid(bytBuf, p, 2)         '' ��
            End With ' MakeDate
        End With ' head
        .KettoNum = IncMid(bytBuf, p, 10)           '' �����o�^�ԍ�
        .DelKubun = IncMid(bytBuf, p, 1)            '' �����n�����敪
        With .RegDate
            .Year = IncMid(bytBuf, p, 4)            '' �N
            .Month = IncMid(bytBuf, p, 2)           '' ��
            .Day = IncMid(bytBuf, p, 2)             '' ��
        End With ' RegDate
        With .DelDate
            .Year = IncMid(bytBuf, p, 4)            '' �N
            .Month = IncMid(bytBuf, p, 2)           '' ��
            .Day = IncMid(bytBuf, p, 2)             '' ��
        End With ' DelDate
        With .BirthDate
            .Year = IncMid(bytBuf, p, 4)            '' �N
            .Month = IncMid(bytBuf, p, 2)           '' ��
            .Day = IncMid(bytBuf, p, 2)             '' ��
        End With ' BirthDate
        .Bamei = IncMid(bytBuf, p, 36)              '' �n��
        .BameiKana = IncMid(bytBuf, p, 36)          '' �n�����p�J�i
        .BameiEng = IncMid(bytBuf, p, 80)           '' �n������
        .UmaKigoCD = IncMid(bytBuf, p, 2)           '' �n�L���R�[�h
        .SexCD = IncMid(bytBuf, p, 1)               '' ���ʃR�[�h
        .HinsyuCD = IncMid(bytBuf, p, 1)            '' �i��R�[�h
        .KeiroCD = IncMid(bytBuf, p, 2)             '' �ѐF�R�[�h
        For i = 0 To 13
            With .Ketto3Info(i)
                .HansyokuNum = IncMid(bytBuf, p, 8) '' �ɐB�o�^�ԍ�
                .Bamei = IncMid(bytBuf, p, 36)      '' �n��
            End With ' Ketto3Info
        Next i
        .TozaiCD = IncMid(bytBuf, p, 1)             '' ���������R�[�h
        .ChokyosiCode = IncMid(bytBuf, p, 5)        '' �����t�R�[�h
        .ChokyosiRyakusyo = IncMid(bytBuf, p, 8)    '' �����t������
        .Syotai = IncMid(bytBuf, p, 20)             '' ���Ғn�於
        .BreederCode = IncMid(bytBuf, p, 6)         '' ���Y�҃R�[�h
        .BreederName = IncMid(bytBuf, p, 70)        '' ���Y�Җ�
        .SanchiName = IncMid(bytBuf, p, 20)         '' �Y�n��
        .BanusiCode = IncMid(bytBuf, p, 6)          '' �n��R�[�h
        .BanusiName = IncMid(bytBuf, p, 64)         '' �n�喼
        .RuikeiHonsyoHeiti = IncMid(bytBuf, p, 9)   '' ���n�{�܋��݌v
        .RuikeiHonsyoSyogai = IncMid(bytBuf, p, 9)  '' ��Q�{�܋��݌v
        .RuikeiFukaHeichi = IncMid(bytBuf, p, 9)    '' ���n�t���܋��݌v
        .RuikeiFukaSyogai = IncMid(bytBuf, p, 9)    '' ��Q�t���܋��݌v
        .RuikeiSyutokuHeichi = IncMid(bytBuf, p, 9) '' ���n�����܋��݌v
        .RuikeiSyutokuSyogai = IncMid(bytBuf, p, 9) '' ��Q�����܋��݌v
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
            .Kyakusitu(i) = IncMid(bytBuf, p, 3)    '' �r���X��
        Next i
        .RaceCount = IncMid(bytBuf, p, 3)           '' �o�^���[�X��
        .crlf = IncMid(bytBuf, p, 2)                '' ���R�[�h��؂�
    End With

    '�o�b�t�@�̈���
    Erase bytBuf
   
    End Sub


    '****** �P�Q�D�R��}�X�^ ****************************************

    Public Sub SetData_KS(lBuf As String, ByRef mBuf As JV_KS_KISYU)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMid(bytBuf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)        '' �N
                .Month = IncMid(bytBuf, p, 2)       '' ��
                .Day = IncMid(bytBuf, p, 2) '' ��
            End With ' MakeDate
        End With ' head
        .KisyuCode = IncMid(bytBuf, p, 5)   '' �R��R�[�h
        .DelKubun = IncMid(bytBuf, p, 1)    '' �R�薕���敪
        With .IssueDate
            .Year = IncMid(bytBuf, p, 4)    '' �N
            .Month = IncMid(bytBuf, p, 2)   '' ��
            .Day = IncMid(bytBuf, p, 2)     '' ��
        End With ' IssueDate
        With .DelDate
            .Year = IncMid(bytBuf, p, 4)    '' �N
            .Month = IncMid(bytBuf, p, 2)   '' ��
            .Day = IncMid(bytBuf, p, 2)     '' ��
        End With ' DelDate
        With .BirthDate
            .Year = IncMid(bytBuf, p, 4)    '' �N
            .Month = IncMid(bytBuf, p, 2)   '' ��
            .Day = IncMid(bytBuf, p, 2)     '' ��
        End With ' BirthDate
        .KisyuName = IncMid(bytBuf, p, 34)  '' �R�薼����
        .reserved = IncMid(bytBuf, p, 34)   '' �\��
        .KisyuNameKana = IncMid(bytBuf, p, 30)      '' �R�薼���p�J�i
        .KisyuRyakusyo = IncMid(bytBuf, p, 8)       '' �R�薼����
        .KisyuNameEng = IncMid(bytBuf, p, 80)       '' �R�薼����
        .SexCD = IncMid(bytBuf, p, 1)       '' ���ʋ敪
        .SikakuCD = IncMid(bytBuf, p, 1)    '' �R�掑�i�R�[�h
        .MinaraiCD = IncMid(bytBuf, p, 1)   '' �R�茩�K�R�[�h
        .TozaiCD = IncMid(bytBuf, p, 1)     '' �R�蓌�������R�[�h
        .Syotai = IncMid(bytBuf, p, 20)     '' ���Ғn�於
        .ChokyosiCode = IncMid(bytBuf, p, 5)        '' ���������t�R�[�h
        .ChokyosiRyakusyo = IncMid(bytBuf, p, 8)    '' ���������t������
        For i = 0 To 1
            With .HatuKiJyo(i)
                With .Hatukijyoid
                    .Year = IncMid(bytBuf, p, 4)    '' �J�ÔN
                    .MonthDay = IncMid(bytBuf, p, 4)        '' �J�Ì���
                    .JyoCD = IncMid(bytBuf, p, 2)   '' ���n��R�[�h
                    .Kaiji = IncMid(bytBuf, p, 2)   '' �J�É�[��N��]
                    .Nichiji = IncMid(bytBuf, p, 2) '' �J�Ó���[N����]
                    .racenum = IncMid(bytBuf, p, 2) '' ���[�X�ԍ�
                End With ' Hatukijyoid
                .SyussoTosu = IncMid(bytBuf, p, 2)  '' �o������
                .KettoNum = IncMid(bytBuf, p, 10)   '' �����o�^�ԍ�
                .Bamei = IncMid(bytBuf, p, 36)      '' �n��
                .KakuteiJyuni = IncMid(bytBuf, p, 2)        '' �m�蒅��
                .IJyoCD = IncMid(bytBuf, p, 1)      '' �ُ�敪�R�[�h
            End With ' HatuKiJyo
        Next i
        For i = 0 To 1
            With .HatuSyori(i)
                With .Hatusyoriid
                    .Year = IncMid(bytBuf, p, 4)    '' �J�ÔN
                    .MonthDay = IncMid(bytBuf, p, 4)        '' �J�Ì���
                    .JyoCD = IncMid(bytBuf, p, 2)   '' ���n��R�[�h
                    .Kaiji = IncMid(bytBuf, p, 2)   '' �J�É�[��N��]
                    .Nichiji = IncMid(bytBuf, p, 2) '' �J�Ó���[N����]
                    .racenum = IncMid(bytBuf, p, 2) '' ���[�X�ԍ�
                End With ' Hatusyoriid
                .SyussoTosu = IncMid(bytBuf, p, 2)  '' �o������
                .KettoNum = IncMid(bytBuf, p, 10)   '' �����o�^�ԍ�
                .Bamei = IncMid(bytBuf, p, 36)      '' �n��
            End With ' HatuSyori
        Next i
        For i = 0 To 2
            With .SaikinJyusyo(i)
                With .SaikinJyusyoid
                    .Year = IncMid(bytBuf, p, 4)    '' �J�ÔN
                    .MonthDay = IncMid(bytBuf, p, 4)        '' �J�Ì���
                    .JyoCD = IncMid(bytBuf, p, 2)   '' ���n��R�[�h
                    .Kaiji = IncMid(bytBuf, p, 2)   '' �J�É�[��N��]
                    .Nichiji = IncMid(bytBuf, p, 2) '' �J�Ó���[N����]
                    .racenum = IncMid(bytBuf, p, 2) '' ���[�X�ԍ�
                End With ' SaikinJyusyoid
                .Hondai = IncMid(bytBuf, p, 60)     '' �������{��
                .Ryakusyo10 = IncMid(bytBuf, p, 20) '' ����������10��
                .Ryakusyo6 = IncMid(bytBuf, p, 12)  '' ����������6��
                .Ryakusyo3 = IncMid(bytBuf, p, 6)   '' ����������3��
                .GradeCD = IncMid(bytBuf, p, 1)     '' �O���[�h�R�[�h
                .SyussoTosu = IncMid(bytBuf, p, 2)  '' �o������
                .KettoNum = IncMid(bytBuf, p, 10)   '' �����o�^�ԍ�
                .Bamei = IncMid(bytBuf, p, 36)      '' �n��
            End With ' SaikinJyusyo
        Next i
        For i = 0 To 2
            With .HonZenRuikei(i)
                .SetYear = IncMid(bytBuf, p, 4)     '' �ݒ�N
                .HonSyokinHeichi = IncMid(bytBuf, p, 10)    '' ���n�{�܋����v
                .HonSyokinSyogai = IncMid(bytBuf, p, 10)    '' ��Q�{�܋����v
                .FukaSyokinHeichi = IncMid(bytBuf, p, 10)   '' ���n�t���܋����v
                .FukaSyokinSyogai = IncMid(bytBuf, p, 10)   '' ��Q�t���܋����v
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
        .crlf = IncMid(bytBuf, p, 2)        '' ���R�[�h��؂�
    End With

    '�o�b�t�@�̈���
    Erase bytBuf
    
    End Sub


    '****** �P�R�D�����t�}�X�^ ****************************************

    Public Sub SetData_CH(lBuf As String, ByRef mBuf As JV_CH_CHOKYOSI)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)              '' ���R�[�h���
            .DataKubun = IncMid(bytBuf, p, 1)               '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)                '' �N
                .Month = IncMid(bytBuf, p, 2)               '' ��
                .Day = IncMid(bytBuf, p, 2)                 '' ��
            End With ' MakeDate
        End With ' head
        .ChokyosiCode = IncMid(bytBuf, p, 5)                '' �����t�R�[�h
        .DelKubun = IncMid(bytBuf, p, 1)                    '' �����t�����敪
        With .IssueDate
            .Year = IncMid(bytBuf, p, 4)                    '' �N
            .Month = IncMid(bytBuf, p, 2)                   '' ��
            .Day = IncMid(bytBuf, p, 2)                     '' ��
        End With ' IssueDate
        With .DelDate
            .Year = IncMid(bytBuf, p, 4)                    '' �N
            .Month = IncMid(bytBuf, p, 2)                   '' ��
            .Day = IncMid(bytBuf, p, 2)                     '' ��
        End With ' DelDate
        With .BirthDate
            .Year = IncMid(bytBuf, p, 4)                    '' �N
            .Month = IncMid(bytBuf, p, 2)                   '' ��
            .Day = IncMid(bytBuf, p, 2)                     '' ��
        End With ' BirthDate
        .ChokyosiName = IncMid(bytBuf, p, 34)               '' �����t������
        .ChokyosiNameKana = IncMid(bytBuf, p, 30)           '' �����t�����p�J�i
        .ChokyosiRyakusyo = IncMid(bytBuf, p, 8)            '' �����t������
        .ChokyosiNameEng = IncMid(bytBuf, p, 80)            '' �����t������
        .SexCD = IncMid(bytBuf, p, 1)                       '' ���ʋ敪
        .TozaiCD = IncMid(bytBuf, p, 1)                     '' �����t���������R�[�h
        .Syotai = IncMid(bytBuf, p, 20)                     '' ���Ғn�於
        For i = 0 To 2
            With .SaikinJyusyo(i)
                With .SaikinJyusyoid
                    .Year = IncMid(bytBuf, p, 4)            '' �J�ÔN
                    .MonthDay = IncMid(bytBuf, p, 4)        '' �J�Ì���
                    .JyoCD = IncMid(bytBuf, p, 2)           '' ���n��R�[�h
                    .Kaiji = IncMid(bytBuf, p, 2)           '' �J�É�[��N��]
                    .Nichiji = IncMid(bytBuf, p, 2)         '' �J�Ó���[N����]
                    .racenum = IncMid(bytBuf, p, 2)         '' ���[�X�ԍ�
                End With ' SaikinJyusyoid
                .Hondai = IncMid(bytBuf, p, 60)             '' �������{��
                .Ryakusyo10 = IncMid(bytBuf, p, 20)         '' ����������10��
                .Ryakusyo6 = IncMid(bytBuf, p, 12)          '' ����������6��
                .Ryakusyo3 = IncMid(bytBuf, p, 6)           '' ����������3��
                .GradeCD = IncMid(bytBuf, p, 1)             '' �O���[�h�R�[�h
                .SyussoTosu = IncMid(bytBuf, p, 2)          '' �o������
                .KettoNum = IncMid(bytBuf, p, 10)           '' �����o�^�ԍ�
                .Bamei = IncMid(bytBuf, p, 36)              '' �n��
            End With ' SaikinJyusyo
        Next i
        For i = 0 To 2
            With .HonZenRuikei(i)
                .SetYear = IncMid(bytBuf, p, 4)             '' �ݒ�N
                .HonSyokinHeichi = IncMid(bytBuf, p, 10)    '' ���n�{�܋����v
                .HonSyokinSyogai = IncMid(bytBuf, p, 10)    '' ��Q�{�܋����v
                .FukaSyokinHeichi = IncMid(bytBuf, p, 10)   '' ���n�t���܋����v
                .FukaSyokinSyogai = IncMid(bytBuf, p, 10)   '' ��Q�t���܋����v
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
        .crlf = IncMid(bytBuf, p, 2)        '' ���R�[�h��؂�
    End With

    '�o�b�t�@�̈���
    Erase bytBuf
   
    End Sub


    '******�P�S�D���Y�҃}�X�^ ****************************************

    Public Sub SetData_BR(lBuf As String, ByRef mBuf As JV_BR_BREEDER)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)              '' ���R�[�h���
            .DataKubun = IncMid(bytBuf, p, 1)               '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)                '' �N
                .Month = IncMid(bytBuf, p, 2)               '' ��
                .Day = IncMid(bytBuf, p, 2)                 '' ��
            End With ' MakeDate
        End With ' head
        .BreederCode = IncMid(bytBuf, p, 6)                 '' ���Y�҃R�[�h
        .BreederName_Co = IncMid(bytBuf, p, 70)             '' ���Y�Җ�(�@�l�i�L�j
        .BreederName = IncMid(bytBuf, p, 70)                '' ���Y�Җ�(�@�l�i���j
        .BreederNameKana = IncMid(bytBuf, p, 70)            '' ���Y�Җ����p�J�i
        .BreederNameEng = IncMid(bytBuf, p, 168)            '' ���Y�Җ�����
        .Address = IncMid(bytBuf, p, 20)                    '' ���Y�ҏZ�������Ȗ�
        For i = 0 To 1
            With .HonRuikei(i)
                .SetYear = IncMid(bytBuf, p, 4)             '' �ݒ�N
                .HonSyokinTotal = IncMid(bytBuf, p, 10)     '' �{�܋����v
                .Fukasyokin = IncMid(bytBuf, p, 10)         '' �t���܋����v
                For j = 0 To 5
                    .Chakukaisu(j) = IncMid(bytBuf, p, 6)   '' ����
                Next j
            End With ' HonRuikei
        Next i
        .crlf = IncMid(bytBuf, p, 2)                        '' ���R�[�h��؂�
    End With

    '�o�b�t�@�̈���
    Erase bytBuf
    
    End Sub


    '****** �P�T�D�n��}�X�^ ****************************************

    Public Sub SetData_BN(lBuf As String, ByRef mBuf As JV_BN_BANUSI)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)              '' ���R�[�h���
            .DataKubun = IncMid(bytBuf, p, 1)               '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)                '' �N
                .Month = IncMid(bytBuf, p, 2)               '' ��
                .Day = IncMid(bytBuf, p, 2)                 '' ��
            End With ' MakeDate
        End With ' head
        .BanusiCode = IncMid(bytBuf, p, 6)                  '' �n��R�[�h
        .BanusiName_Co = IncMid(bytBuf, p, 64)              '' �n�喼�i�@�l�i�L�j
        .BanusiName = IncMid(bytBuf, p, 64)                 '' �n�喼�i�@�l�i���j
        .BanusiNameKana = IncMid(bytBuf, p, 50)             '' �n�喼���p�J�i
        .BanusiNameEng = IncMid(bytBuf, p, 100)             '' �n�喼����
        .Fukusyoku = IncMid(bytBuf, p, 60)                  '' ���F�W��
        For i = 0 To 1
            With .HonRuikei(i)
                .SetYear = IncMid(bytBuf, p, 4)             '' �ݒ�N
                .HonSyokinTotal = IncMid(bytBuf, p, 10)     '' �{�܋����v
                .Fukasyokin = IncMid(bytBuf, p, 10)         '' �t���܋����v
                For j = 0 To 5
                    .Chakukaisu(j) = IncMid(bytBuf, p, 6)   '' ����
                Next j
            End With ' HonRuikei
        Next i
        .crlf = IncMid(bytBuf, p, 2)                        '' ���R�[�h��؂�
    End With

    '�o�b�t�@�̈���
    Erase bytBuf
   
    End Sub

    '****** �P�U�D�ɐB�n�}�X�^ ****************************************

    Public Sub SetData_HN(lBuf As String, ByRef mBuf As JV_HN_HANSYOKU)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMid(bytBuf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)        '' �N
                .Month = IncMid(bytBuf, p, 2)       '' ��
                .Day = IncMid(bytBuf, p, 2)         '' ��
            End With ' MakeDate
        End With ' head
        .HansyokuNum = IncMid(bytBuf, p, 8)         '' �ɐB�o�^�ԍ�
        .reserved = IncMid(bytBuf, p, 8)            '' �\��
        .KettoNum = IncMid(bytBuf, p, 10)           '' �����o�^�ԍ�
        .DelKubun = IncMid(bytBuf, p, 1)            '' �ɐB�n�����敪
        .Bamei = IncMid(bytBuf, p, 36)              '' �n��
        .BameiKana = IncMid(bytBuf, p, 40)          '' �n�����p�J�i
        .BameiEng = IncMid(bytBuf, p, 80)           '' �n������
        .BirthYear = IncMid(bytBuf, p, 4)           '' ���N
        .SexCD = IncMid(bytBuf, p, 1)               '' ���ʃR�[�h
        .HinsyuCD = IncMid(bytBuf, p, 1)            '' �i��R�[�h
        .KeiroCD = IncMid(bytBuf, p, 2)             '' �ѐF�R�[�h
        .HansyokuMochiKubun = IncMid(bytBuf, p, 1)  '' �ɐB�n�����敪
        .ImportYear = IncMid(bytBuf, p, 4)          '' �A���N
        .SanchiName = IncMid(bytBuf, p, 20)         '' �Y�n��
        .HansyokuFNum = IncMid(bytBuf, p, 8)        '' ���n�ɐB�o�^�ԍ�
        .HansyokuMNum = IncMid(bytBuf, p, 8)        '' ��n�ɐB�o�^�ԍ�
        .crlf = IncMid(bytBuf, p, 2)                '' ���R�[�h��؂�
    End With

    '�o�b�t�@�̈���
    Erase bytBuf
   
    End Sub


    '****** �P�V�D�Y��}�X�^ ****************************************

    Public Sub SetData_SK(lBuf As String, ByRef mBuf As JV_SK_SANKU)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMid(bytBuf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)        '' �N
                .Month = IncMid(bytBuf, p, 2)       '' ��
                .Day = IncMid(bytBuf, p, 2)         '' ��
            End With ' MakeDate
        End With ' head
        .KettoNum = IncMid(bytBuf, p, 10)           '' �����o�^�ԍ�
        With .BirthDate
            .Year = IncMid(bytBuf, p, 4)            '' �N
            .Month = IncMid(bytBuf, p, 2)           '' ��
            .Day = IncMid(bytBuf, p, 2)             '' ��
        End With ' BirthDate
        .SexCD = IncMid(bytBuf, p, 1)               '' ���ʃR�[�h
        .HinsyuCD = IncMid(bytBuf, p, 1)            '' �i��R�[�h
        .KeiroCD = IncMid(bytBuf, p, 2)             '' �ѐF�R�[�h
        .SankuMochiKubun = IncMid(bytBuf, p, 1)     '' �Y����敪
        .ImportYear = IncMid(bytBuf, p, 4)          '' �A���N
        .BreederCode = IncMid(bytBuf, p, 6)         '' ���Y�҃R�[�h
        .SanchiName = IncMid(bytBuf, p, 20)         '' �Y�n��
        For i = 0 To 13
            .HansyokuNum(i) = IncMid(bytBuf, p, 8)  '' 3�㌌��
        Next i
    End With

    '�o�b�t�@�̈���
    Erase bytBuf
    
    End Sub

    '****** �P�W�D���R�[�h�}�X�^ ****************************************

    Public Sub SetData_RC(lBuf As String, ByRef mBuf As JV_RC_RECORD)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)              '' ���R�[�h���
            .DataKubun = IncMid(bytBuf, p, 1)               '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)                '' �N
                .Month = IncMid(bytBuf, p, 2)               '' ��
                .Day = IncMid(bytBuf, p, 2)                 '' ��
            End With ' MakeDate
        End With ' head
        .RecInfoKubun = IncMid(bytBuf, p, 1)                '' ���R�[�h���ʋ敪
        With .id
            .Year = IncMid(bytBuf, p, 4)                    '' �J�ÔN
            .MonthDay = IncMid(bytBuf, p, 4)                '' �J�Ì���
            .JyoCD = IncMid(bytBuf, p, 2)                   '' ���n��R�[�h
            .Kaiji = IncMid(bytBuf, p, 2)                   '' �J�É�[��N��]
            .Nichiji = IncMid(bytBuf, p, 2)                 '' �J�Ó���[N����]
            .racenum = IncMid(bytBuf, p, 2)                 '' ���[�X�ԍ�
        End With ' id
        .TokuNum = IncMid(bytBuf, p, 4)                     '' ���ʋ����ԍ�
        .Hondai = IncMid(bytBuf, p, 60)                     '' �������{��
        .GradeCD = IncMid(bytBuf, p, 1)                     '' �O���[�h�R�[�h
        .SyubetuCD = IncMid(bytBuf, p, 2)                   '' ������ʃR�[�h
        .Kyori = IncMid(bytBuf, p, 4)                       '' ����
        .TrackCD = IncMid(bytBuf, p, 2)                     '' �g���b�N�R�[�h
        .RecKubun = IncMid(bytBuf, p, 1)                    '' ���R�[�h�敪
        .RecTime = IncMid(bytBuf, p, 4)                     '' ���R�[�h�^�C��
        With .TenkoBaba
            .TenkoCD = IncMid(bytBuf, p, 1)                 '' �V��R�[�h
            .SibaBabaCD = IncMid(bytBuf, p, 1)              '' �Ŕn���ԃR�[�h
            .DirtBabaCD = IncMid(bytBuf, p, 1)              '' �_�[�g�n���ԃR�[�h
        End With ' TenkoBaba
        For i = 0 To 2
            With .RecUmaInfo(i)
                .KettoNum = IncMid(bytBuf, p, 10)           '' �����o�^�ԍ�
                .Bamei = IncMid(bytBuf, p, 36)              '' �n��
                .UmaKigoCD = IncMid(bytBuf, p, 2)           '' �n�L���R�[�h
                .SexCD = IncMid(bytBuf, p, 1)               '' ���ʃR�[�h
                .ChokyosiCode = IncMid(bytBuf, p, 5)        '' �����t�R�[�h
                .ChokyosiName = IncMid(bytBuf, p, 34)       '' �����t��
                .Futan = IncMid(bytBuf, p, 3)               '' ���S�d��
                .KisyuCode = IncMid(bytBuf, p, 5)           '' �R��R�[�h
                .KisyuName = IncMid(bytBuf, p, 34)          '' �R�薼
            End With ' RecUmaInfo
        Next i
        .crlf = IncMid(bytBuf, p, 2)                        '' ���R�[�h��؂�
    End With

    '�o�b�t�@�̈���
    Erase bytBuf

    End Sub


    '****** �P�X�D��H���� ****************************************

    Public Sub SetData_HC(lBuf As String, ByRef mBuf As JV_HC_HANRO)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)  '' ���R�[�h���
            .DataKubun = IncMid(bytBuf, p, 1)   '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)    '' �N
                .Month = IncMid(bytBuf, p, 2)   '' ��
                .Day = IncMid(bytBuf, p, 2)     '' ��
            End With ' MakeDate
        End With ' head
        .TresenKubun = IncMid(bytBuf, p, 1)     '' �g���Z���敪
        With .ChokyoDate
            .Year = IncMid(bytBuf, p, 4)        '' �N
            .Month = IncMid(bytBuf, p, 2)       '' ��
            .Day = IncMid(bytBuf, p, 2)         '' ��
        End With ' ChokyoDate
        .ChokyoTime = IncMid(bytBuf, p, 4)      '' ��������
        .KettoNum = IncMid(bytBuf, p, 10)       '' �����o�^�ԍ�
        .HaronTime4 = IncMid(bytBuf, p, 4)      '' 4�n�����^�C�����v(800M-0M)
        .LapTime4 = IncMid(bytBuf, p, 3)        '' ���b�v�^�C��(800M-600M)
        .HaronTime3 = IncMid(bytBuf, p, 4)      '' 3�n�����^�C�����v(600M-0M)
        .LapTime3 = IncMid(bytBuf, p, 3)        '' ���b�v�^�C��(600M-400M)
        .HaronTime2 = IncMid(bytBuf, p, 4)      '' 2�n�����^�C�����v(400M-0M)
        .LapTime2 = IncMid(bytBuf, p, 3)        '' ���b�v�^�C��(400M-200M)
        .LapTime1 = IncMid(bytBuf, p, 3)        '' ���b�v�^�C��(200M-0M)
        .crlf = IncMid(bytBuf, p, 2)            '' ���R�[�h��؂�
    End With

    '�o�b�t�@�̈���
    Erase bytBuf
    
  End Sub


    '****** �Q�O�D�n�̏d ****************************************

    Public Sub SetData_WH(lBuf As String, ByRef mBuf As JV_WH_BATAIJYU)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMid(bytBuf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)        '' �N
                .Month = IncMid(bytBuf, p, 2)       '' ��
                .Day = IncMid(bytBuf, p, 2)         '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)            '' �J�ÔN
            .MonthDay = IncMid(bytBuf, p, 4)        '' �J�Ì���
            .JyoCD = IncMid(bytBuf, p, 2)           '' ���n��R�[�h
            .Kaiji = IncMid(bytBuf, p, 2)           '' �J�É�[��N��]
            .Nichiji = IncMid(bytBuf, p, 2)         '' �J�Ó���[N����]
            .racenum = IncMid(bytBuf, p, 2)         '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = IncMid(bytBuf, p, 2)           '' ��
            .Day = IncMid(bytBuf, p, 2)             '' ��
            .Hour = IncMid(bytBuf, p, 2)            '' ��
            .Minute = IncMid(bytBuf, p, 2)          '' ��
        End With ' HappyoTime
        For i = 0 To 17
            With .BataijyuInfo(i)
                .Umaban = IncMid(bytBuf, p, 2)      '' �n��
                .Bamei = IncMid(bytBuf, p, 36)      '' �n��
                .BaTaijyu = IncMid(bytBuf, p, 3)    '' �n�̏d
                .ZogenFugo = IncMid(bytBuf, p, 1)   '' ��������
                .ZogenSa = IncMid(bytBuf, p, 3)     '' ������
            End With ' BataijyuInfo
        Next i
        .crlf = IncMid(bytBuf, p, 2)                '' ���R�[�h��؂�
    End With

    '�o�b�t�@�̈���
    Erase bytBuf
   
    End Sub


    '****** �Q�P�D�V��n���� ******************************************

    Public Sub SetData_WE(lBuf As String, ByRef mBuf As JV_WE_WEATHER)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMid(bytBuf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)        '' �N
                .Month = IncMid(bytBuf, p, 2)       '' ��
                .Day = IncMid(bytBuf, p, 2)         '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)            '' �J�ÔN
            .MonthDay = IncMid(bytBuf, p, 4)        '' �J�Ì���
            .JyoCD = IncMid(bytBuf, p, 2)           '' ���n��R�[�h
            .Kaiji = IncMid(bytBuf, p, 2)           '' �J�É�[��N��]
            .Nichiji = IncMid(bytBuf, p, 2)         '' �J�Ó���[N����]
        End With ' id
        With .HappyoTime
            .Month = IncMid(bytBuf, p, 2)           '' ��
            .Day = IncMid(bytBuf, p, 2)             '' ��
            .Hour = IncMid(bytBuf, p, 2)            '' ��
            .Minute = IncMid(bytBuf, p, 2)          '' ��
        End With ' HappyoTime
        .HenkoID = IncMid(bytBuf, p, 1)             '' �ύX����
        With .TenkoBaba
            .TenkoCD = IncMid(bytBuf, p, 1)         '' �V��R�[�h
            .SibaBabaCD = IncMid(bytBuf, p, 1)      '' �Ŕn���ԃR�[�h
            .DirtBabaCD = IncMid(bytBuf, p, 1)      '' �_�[�g�n���ԃR�[�h
        End With ' TenkoBaba
        With .TenkoBabaBefore
            .TenkoCD = IncMid(bytBuf, p, 1)         '' �V��R�[�h
            .SibaBabaCD = IncMid(bytBuf, p, 1)      '' �Ŕn���ԃR�[�h
            .DirtBabaCD = IncMid(bytBuf, p, 1)      '' �_�[�g�n���ԃR�[�h
        End With ' TenkoBabaBefore
        .crlf = IncMid(bytBuf, p, 2)                '' ���R�[�h��؂�
    End With

    '�o�b�t�@�̈���
    Erase bytBuf
    
    End Sub


    '****** �Q�Q�D�o������E�������O ****************************************

    Public Sub SetData_AV(lBuf As String, ByRef mBuf As JV_AV_INFO)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMid(bytBuf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)        '' �N
                .Month = IncMid(bytBuf, p, 2)       '' ��
                .Day = IncMid(bytBuf, p, 2)         '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)            '' �J�ÔN
            .MonthDay = IncMid(bytBuf, p, 4)        '' �J�Ì���
            .JyoCD = IncMid(bytBuf, p, 2)           '' ���n��R�[�h
            .Kaiji = IncMid(bytBuf, p, 2)           '' �J�É�[��N��]
            .Nichiji = IncMid(bytBuf, p, 2)         '' �J�Ó���[N����]
            .racenum = IncMid(bytBuf, p, 2)         '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = IncMid(bytBuf, p, 2)           '' ��
            .Day = IncMid(bytBuf, p, 2)             '' ��
            .Hour = IncMid(bytBuf, p, 2)            '' ��
            .Minute = IncMid(bytBuf, p, 2)          '' ��
        End With ' HappyoTime
        .Umaban = IncMid(bytBuf, p, 2)              '' �n��
        .Bamei = IncMid(bytBuf, p, 36)              '' �n��
        .JiyuKubun = IncMid(bytBuf, p, 3)           '' ���R�敪
        .crlf = IncMid(bytBuf, p, 2)                '' ���R�[�h��؂�
    End With

    '�o�b�t�@�̈���
    Erase bytBuf
    
    End Sub

    '************ �Q�R�D�R��ύX ****************************************
  
    Public Sub SetData_JC(lBuf As String, ByRef mBuf As JV_JC_INFO)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)  '' ���R�[�h���
            .DataKubun = IncMid(bytBuf, p, 1)   '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)    '' �N
                .Month = IncMid(bytBuf, p, 2)   '' ��
                .Day = IncMid(bytBuf, p, 2)     '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)        '' �J�ÔN
            .MonthDay = IncMid(bytBuf, p, 4)        '' �J�Ì���
            .JyoCD = IncMid(bytBuf, p, 2)       '' ���n��R�[�h
            .Kaiji = IncMid(bytBuf, p, 2)       '' �J�É�[��N��]
            .Nichiji = IncMid(bytBuf, p, 2)     '' �J�Ó���[N����]
            .racenum = IncMid(bytBuf, p, 2)     '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = IncMid(bytBuf, p, 2)       '' ��
            .Day = IncMid(bytBuf, p, 2)         '' ��
            .Hour = IncMid(bytBuf, p, 2)        '' ��
            .Minute = IncMid(bytBuf, p, 2)      '' ��
        End With ' HappyoTime
        .Umaban = IncMid(bytBuf, p, 2)          '' �n��
        .Bamei = IncMid(bytBuf, p, 36)          '' �n��
        With .JCInfoAfter
            .Futan = IncMid(bytBuf, p, 3)       '' ���S�d��
            .KisyuCode = IncMid(bytBuf, p, 5)   '' �R��R�[�h
            .KisyuName = IncMid(bytBuf, p, 34)  '' �R�薼
            .MinaraiCD = IncMid(bytBuf, p, 1)   '' �R�茩�K�R�[�h
        End With ' JCInfoAfter
        With .JCInfoBefore
            .Futan = IncMid(bytBuf, p, 3)       '' ���S�d��
            .KisyuCode = IncMid(bytBuf, p, 5)   '' �R��R�[�h
            .KisyuName = IncMid(bytBuf, p, 34)  '' �R�薼
            .MinaraiCD = IncMid(bytBuf, p, 1)   '' �R�茩�K�R�[�h
        End With ' JCInfoBefore
        .crlf = IncMid(bytBuf, p, 2)            '' ���R�[�h��؂�
    End With

    '�o�b�t�@�̈���
    Erase bytBuf
   
    End Sub

    '****** �Q�S�D�f�[�^�}�C�j���O�\�z***********************************
    
    Public Sub SetData_DM(lBuf As String, ByRef mBuf As JV_DM_INFO)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)  '' ���R�[�h���
            .DataKubun = IncMid(bytBuf, p, 1)   '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)    '' �N
                .Month = IncMid(bytBuf, p, 2)   '' ��
                .Day = IncMid(bytBuf, p, 2)     '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)        '' �J�ÔN
            .MonthDay = IncMid(bytBuf, p, 4)    '' �J�Ì���
            .JyoCD = IncMid(bytBuf, p, 2)       '' ���n��R�[�h
            .Kaiji = IncMid(bytBuf, p, 2)       '' �J�É�[��N��]
            .Nichiji = IncMid(bytBuf, p, 2)     '' �J�Ó���[N����]
            .racenum = IncMid(bytBuf, p, 2)     '' ���[�X�ԍ�
        End With ' id
        With .MakeHM
            .Hour = IncMid(bytBuf, p, 2)        '' ��
            .Minute = IncMid(bytBuf, p, 2)      '' ��
        End With ' MakeHM
        For i = 0 To 17
            With .DMInfo(i)
                .Umaban = IncMid(bytBuf, p, 2)  '' �n��
                .DMTime = IncMid(bytBuf, p, 5)  '' �\�z���j�^�C��
                .DMGosaP = IncMid(bytBuf, p, 4) '' �\�z�덷(�M���x)�{
                .DMGosaM = IncMid(bytBuf, p, 4) '' �\�z�덷(�M���x)�|
            End With ' DMInfo
        Next i
        .crlf = IncMid(bytBuf, p, 2)            '' ���R�[�h��؂�
    End With

    '�o�b�t�@�̈���
    Erase bytBuf
    
    End Sub


    '****** �Q�T�D�J�ÃX�P�W���[��************************************
    
    Public Sub SetData_YS(lBuf As String, ByRef mBuf As JV_YS_SCHEDULE)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMid(bytBuf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)        '' �N
                .Month = IncMid(bytBuf, p, 2)       '' ��
                .Day = IncMid(bytBuf, p, 2)         '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)            '' �J�ÔN
            .MonthDay = IncMid(bytBuf, p, 4)        '' �J�Ì���
            .JyoCD = IncMid(bytBuf, p, 2)           '' ���n��R�[�h
            .Kaiji = IncMid(bytBuf, p, 2)           '' �J�É�[��N��]
            .Nichiji = IncMid(bytBuf, p, 2)         '' �J�Ó���[N����]
        End With ' id
        .YoubiCD = IncMid(bytBuf, p, 1)             '' �j���R�[�h
        For i = 0 To 2
            With .JyusyoInfo(i)
                .TokuNum = IncMid(bytBuf, p, 4)     '' ���ʋ����ԍ�
                .Hondai = IncMid(bytBuf, p, 60)     '' �������{��
                .Ryakusyo10 = IncMid(bytBuf, p, 20) '' ����������10��
                .Ryakusyo6 = IncMid(bytBuf, p, 12)  '' ����������6��
                .Ryakusyo3 = IncMid(bytBuf, p, 6)   '' ����������3��
                .Nkai = IncMid(bytBuf, p, 3)        '' �d�܉�[��N��]
                .GradeCD = IncMid(bytBuf, p, 1)     '' �O���[�h�R�[�h
                .SyubetuCD = IncMid(bytBuf, p, 2)   '' ������ʃR�[�h
                .KigoCD = IncMid(bytBuf, p, 3)      '' �����L���R�[�h
                .JyuryoCD = IncMid(bytBuf, p, 1)    '' �d�ʎ�ʃR�[�h
                .Kyori = IncMid(bytBuf, p, 4)       '' ����
                .TrackCD = IncMid(bytBuf, p, 2)     '' �g���b�N�R�[�h
            End With ' JyusyoInfo
        Next i
        .crlf = IncMid(bytBuf, p, 2)                '' ���R�[�h��؂�
    End With

    '�o�b�t�@�̈���
    Erase bytBuf
    
    End Sub

    Public Sub SetData_H6(lBuf As String, ByRef mBuf As JV_H6_HYOSU_SANRENTAN)
    Dim bytBuf() As Byte                                    '' �o�C�g�z��ŏ������邽�߂̃o�b�t�@
    Dim i As Integer                                        '' ���[�v�J�E���^
    Dim j As Integer                                        '' ���[�v�J�E���^
    Dim k As Integer                                        '' ���[�v�J�E���^
    Dim p As Long                                           '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBuf, vbFromUnicode)
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)              '' ���R�[�h���
            .DataKubun = IncMid(bytBuf, p, 1)               '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)                '' �N
                .Month = IncMid(bytBuf, p, 2)               '' ��
                .Day = IncMid(bytBuf, p, 2)                 '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)                    '' �J�ÔN
            .MonthDay = IncMid(bytBuf, p, 4)                '' �J�Ì���
            .JyoCD = IncMid(bytBuf, p, 2)                   '' ���n��R�[�h
            .Kaiji = IncMid(bytBuf, p, 2)                   '' �J�É�[��N��]
            .Nichiji = IncMid(bytBuf, p, 2)                 '' �J�Ó���[N����]
            .racenum = IncMid(bytBuf, p, 2)                 '' ���[�X�ԍ�
        End With ' id
        .TorokuTosu = IncMid(bytBuf, p, 2)                  '' �o�^����
        .SyussoTosu = IncMid(bytBuf, p, 2)                  '' �o������
        .HatubaiFlag = IncMid(bytBuf, p, 1)                     '' �����t���O
        For i = 0 To 17
            .HenkanUma(i) = IncMid(bytBuf, p, 1)            '' �ԊҔn�ԏ��(�n��01�`18)
        Next i

        For i = 0 To 4895
            With .HyoSanrentan(i)
                .Kumi = IncMid(bytBuf, p, 6)                '' �g��
                .Hyo = IncMid(bytBuf, p, 11)                '' �[��
                .Ninki = IncMid(bytBuf, p, 4)               '' �l�C
            End With ' HyoSanrentan
        Next i
        For i = 0 To 1
            .HyoTotal(i) = IncMid(bytBuf, p, 11)            '' �[�����v
        Next i
        .crlf = IncMid(bytBuf, p, 2)                        '' ���R�[�h��؂�
    End With
    
    '�o�b�t�@�̈���
    Erase bytBuf
    
    End Sub

    Public Sub SetData_O6(lBuf As String, ByRef mBuf As JV_O6_ODDS_SANRENTAN)

    Dim bytBuf() As Byte                                    '' �o�C�g�z��ŏ������邽�߂̃o�b�t�@
    Dim i As Integer                                        '' ���[�v�J�E���^
    Dim j As Integer                                        '' ���[�v�J�E���^
    Dim k As Integer                                        '' ���[�v�J�E���^
    Dim p As Long                                           '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBuf, vbFromUnicode)

    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)              '' ���R�[�h���
            .DataKubun = IncMid(bytBuf, p, 1)               '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)                '' �N
                .Month = IncMid(bytBuf, p, 2)               '' ��
                .Day = IncMid(bytBuf, p, 2)                 '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)                    '' �J�ÔN
            .MonthDay = IncMid(bytBuf, p, 4)                '' �J�Ì���
            .JyoCD = IncMid(bytBuf, p, 2)                   '' ���n��R�[�h
            .Kaiji = IncMid(bytBuf, p, 2)                   '' �J�É�[��N��]
            .Nichiji = IncMid(bytBuf, p, 2)                 '' �J�Ó���[N����]
            .racenum = IncMid(bytBuf, p, 2)                 '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = IncMid(bytBuf, p, 2)                   '' ��
            .Day = IncMid(bytBuf, p, 2)                     '' ��
            .Hour = IncMid(bytBuf, p, 2)                    '' ��
            .Minute = IncMid(bytBuf, p, 2)                  '' ��
        End With ' HappyoTime
        .TorokuTosu = IncMid(bytBuf, p, 2)                  '' �o�^����
        .SyussoTosu = IncMid(bytBuf, p, 2)                  '' �o������
        .SanrentanFlag = IncMid(bytBuf, p, 1)               '' �����t���O�@3�A�P
        For i = 0 To 4895
            With .OddsSanrentanInfo(i)
                .Kumi = IncMid(bytBuf, p, 6)                '' �g��
                .Odds = IncMid(bytBuf, p, 7)                '' �I�b�Y
                .Ninki = IncMid(bytBuf, p, 4)               '' �l�C��
            End With ' OddsSanrentanInfo
        Next i
        .TotalHyosuSanrentan = IncMid(bytBuf, p, 11)        '' 3�A�P�[�����v
        .crlf = IncMid(bytBuf, p, 2)                        '' ���R�[�h��؂�
    End With
    
    '�o�b�t�@�̈���
    Erase bytBuf
    
    End Sub

Public Sub SetData_O6Z(lBuf As String, ByRef mBuf As JV_O6_ODDS_SANRENTAN2)

    Dim bytBuf() As Byte                                    '' �o�C�g�z��ŏ������邽�߂̃o�b�t�@
    Dim i As Integer                                        '' ���[�v�J�E���^
    Dim j As Integer                                        '' ���[�v�J�E���^
    Dim k As Integer                                        '' ���[�v�J�E���^
    Dim p As Long                                           '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBuf, vbFromUnicode)

    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)              '' ���R�[�h���
            .DataKubun = IncMid(bytBuf, p, 1)               '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)                '' �N
                .Month = IncMid(bytBuf, p, 2)               '' ��
                .Day = IncMid(bytBuf, p, 2)                 '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)                    '' �J�ÔN
            .MonthDay = IncMid(bytBuf, p, 4)                '' �J�Ì���
            .JyoCD = IncMid(bytBuf, p, 2)                   '' ���n��R�[�h
            .Kaiji = IncMid(bytBuf, p, 2)                   '' �J�É�[��N��]
            .Nichiji = IncMid(bytBuf, p, 2)                 '' �J�Ó���[N����]
            .racenum = IncMid(bytBuf, p, 2)                 '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = IncMid(bytBuf, p, 2)                   '' ��
            .Day = IncMid(bytBuf, p, 2)                     '' ��
            .Hour = IncMid(bytBuf, p, 2)                    '' ��
            .Minute = IncMid(bytBuf, p, 2)                  '' ��
        End With ' HappyoTime
        .TorokuTosu = IncMid(bytBuf, p, 2)                  '' �o�^����
        .SyussoTosu = IncMid(bytBuf, p, 2)                  '' �o������
        .SanrentanFlag = IncMid(bytBuf, p, 1)               '' �����t���O�@3�A�P
        Set .OddsSanrentanInfo = New Collection
        For i = 0 To 4895
            Set cOddssanrentaninfo = New cODDS_SANRENTAN_INFO
            cOddssanrentaninfo.Kumi = IncMid(bytBuf, p, 6)
            cOddssanrentaninfo.Odds = IncMid(bytBuf, p, 7)
            cOddssanrentaninfo.Ninki = IncMid(bytBuf, p, 4)
            .OddsSanrentanInfo.Add cOddssanrentaninfo
        Next i
'        Debug.Print .OddsSanrentanInfo.Count
        .TotalHyosuSanrentan = IncMid(bytBuf, p, 11)        '' 3�A�P�[�����v
        .crlf = IncMid(bytBuf, p, 2)                        '' ���R�[�h��؂�
    End With
    
    '�o�b�t�@�̈���
    Erase bytBuf
    
    End Sub

    Public Sub SetData_CC(lBuf As String, ByRef mBuf As JV_CC_INFO)

    Dim bytBuf() As Byte                                    '' �o�C�g�z��ŏ������邽�߂̃o�b�t�@
    Dim i As Integer                                        '' ���[�v�J�E���^
    Dim j As Integer                                        '' ���[�v�J�E���^
    Dim k As Integer                                        '' ���[�v�J�E���^
    Dim p As Long                                           '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBuf, vbFromUnicode)

    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)              '' ���R�[�h���
            .DataKubun = IncMid(bytBuf, p, 1)               '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)                '' �N
                .Month = IncMid(bytBuf, p, 2)               '' ��
                .Day = IncMid(bytBuf, p, 2)                 '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)                    '' �J�ÔN
            .MonthDay = IncMid(bytBuf, p, 4)                    '' �J�Ì���
            .JyoCD = IncMid(bytBuf, p, 2)                   '' ���n��R�[�h
            .Kaiji = IncMid(bytBuf, p, 2)                   '' �J�É�[��N��]
            .Nichiji = IncMid(bytBuf, p, 2)                 '' �J�Ó���[N����]
            .racenum = IncMid(bytBuf, p, 2)                 '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = IncMid(bytBuf, p, 2)                   '' ��
            .Day = IncMid(bytBuf, p, 2)                     '' ��
            .Hour = IncMid(bytBuf, p, 2)                    '' ��
            .Minute = IncMid(bytBuf, p, 2)                  '' ��
        End With ' HappyoTime
        
        With .CCInfoAfter
            .Kyori = IncMid(bytBuf, p, 4)                   '' ����
            .TruckCD = IncMid(bytBuf, p, 2)                 '' �g���b�N�R�[�h
        End With ' CCInfoAfter
        With .CCInfoBefore
            .Kyori = IncMid(bytBuf, p, 4)                   '' ����
            .TruckCD = IncMid(bytBuf, p, 2)                 '' �g���b�N�R�[�h
        End With ' CCInfoBefore
        .JiyuCD = IncMid(bytBuf, p, 1)                      '' ���R�R�[�h

        .crlf = IncMid(bytBuf, p, 2)                        '' ���R�[�h��؂�
    End With
    
    '�o�b�t�@�̈���
    Erase bytBuf
    
    End Sub

    Public Sub SetData_TC(lBuf As String, ByRef mBuf As JV_TC_INFO)
    Dim bytBuf() As Byte                                    '' �o�C�g�z��ŏ������邽�߂̃o�b�t�@
    Dim i As Integer                                        '' ���[�v�J�E���^
    Dim j As Integer                                        '' ���[�v�J�E���^
    Dim k As Integer                                        '' ���[�v�J�E���^
    Dim p As Long                                           '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBuf, vbFromUnicode)

    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(bytBuf, p, 2)              '' ���R�[�h���
            .DataKubun = IncMid(bytBuf, p, 1)               '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(bytBuf, p, 4)                '' �N
                .Month = IncMid(bytBuf, p, 2)               '' ��
                .Day = IncMid(bytBuf, p, 2)                 '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(bytBuf, p, 4)                    '' �J�ÔN
            .MonthDay = IncMid(bytBuf, p, 4)                    '' �J�Ì���
            .JyoCD = IncMid(bytBuf, p, 2)                   '' ���n��R�[�h
            .Kaiji = IncMid(bytBuf, p, 2)                   '' �J�É�[��N��]
            .Nichiji = IncMid(bytBuf, p, 2)                 '' �J�Ó���[N����]
            .racenum = IncMid(bytBuf, p, 2)                 '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = IncMid(bytBuf, p, 2)                   '' ��
            .Day = IncMid(bytBuf, p, 2)                     '' ��
            .Hour = IncMid(bytBuf, p, 2)                    '' ��
            .Minute = IncMid(bytBuf, p, 2)                  '' ��
        End With ' HappyoTime
        With .TCInfoAfter
            .Ji = IncMid(bytBuf, p, 2)                          '' ��
            .Fun = IncMid(bytBuf, p, 2)                         '' ��
        End With ' TCInfoAfter
        With .TCInfoBefore
            .Ji = IncMid(bytBuf, p, 2)                          '' ��
            .Fun = IncMid(bytBuf, p, 2)                         '' ��
        End With ' TCInfoBefore

        .crlf = IncMid(bytBuf, p, 2)                        '' ���R�[�h��؂�
    End With
    
    '�o�b�t�@�̈���
    Erase bytBuf
    
    End Sub

 '------------------------------------------------------------------------
 '�@�@�o�C�g�z����o�C�g���Ő؏o��
 '------------------------------------------------------------------------
 Public Function IncMid(ByRef vBuf() As Byte, p As Long, length As Long) As String
     IncMid = StrConv(MidB(vBuf, p, length), vbUnicode)
     p = p + length
 End Function
     
        
