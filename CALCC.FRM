       =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   900
      _StockProps     =   78
      Caption         =   "     &OK"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      AutoSize        =   2
      Picture         =   "CALCC.frx":030A
   End
   Begin Threed.SSCommand cmdAnnuler 
      Cancel          =   -1  'True
      Height          =   510
      Left            =   2280
      TabIndex        =   15
      Top             =   1800
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   900
      _StockProps     =   78
      Caption         =   "        A&nnuler"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   2
      Picture         =   "CALCC.frx":0866
   End
   Begin Threed.SSCommand cmdAide 
      Height          =   510
      Left            =   4440
      TabIndex        =   16
      Top             =   1800
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   900
      _StockProps     =   78
      Caption         =   "    &Aide"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   2
      Picture         =   "CALCC.frx":0E12
   End
End
Attribute VB_Name = "frmCalcC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAide_Click()
    Msg$ = "    Pour conna�tre le code couleur correspondant � une valeur, entrez d'abord cette valeur, puis indiquez si cette valeur est exprim�e en ohm, kiloohm ou m�gohm."
    Msg$ = Msg$ & Enter & "    Indiquer aussi la tol�rance."
    MsgBox Msg$
End Sub

Private Sub cmdAide_GotFocus()
    cmdAide.FontBold = True
End Sub

Private Sub cmdAide_LostFocus()
    cmdAide.FontBold = False
End Sub

Private Sub cmdAnnuler_Click()
    Screen.MousePointer = 11
    DoEvents
    txtValeur.Text = ""
    frmCalcC.Hide
    Screen.MousePointer = 0
End Sub

Private Sub cmdAnnuler_GotFocus()
    cmdAnnuler.FontBold = True
End Sub

Private Sub cmdAnnuler_LostFocus()
    cmdAnnuler.FontBold = False
End Sub

Private Sub cmdOK_Click()
Dim A(3) As Single
    Screen.MousePointer = 11
    
    'Calcule la valeur mais pas la tol�rance
    For i% = 0 To 6 Step 3
        If optOhm(i%).Value Then ValeurEntre = Int(Val(txtValeur.Text) * 10 ^ i%)
    Next
    ValeurTexte = Trim(Str(ValeurEntre))
            
    'v�rifit si elle est juste
    If ValeurEntre < 1 Or ValeurEntre > 99000000 Then
        MsgBox "Vous devez entrer une valeur!" & _
            Enter & "Entre 1 et 99.000.000 ohm.", vbInformation
        cmdOK.Enabled = False
        txtValeur.Text = ""
        txtValeur.SetFocus
        Exit Sub
    End If
    
    Rem Donne A(1), A(2) et A(3) en fonction de ValeurEntre
    If Len(ValeurEntre) = 1 Then
        A(1) = 0
        A(2) = ValeurEntre
       @�:��������6������������ )��@9��w���g� t
�>� t�7��� ÿ(ӹ" 3��
ӣ��PSR� �(Z[X.�.��<Ar<Zw �W����
�t������_�VW2�6�6Y�e�}:u����
�u�
�u��!A�Ћ�N����6�_�  ���6�_��e�s�6�_�:��ª6�_���6�>Y��_^�6�>�� uPVR2�6�>f�:u6�e��߀�@�X�6����!
�u��!����@�U��D:\6�>]� t
���`q���!Z^�SWV�e�U���
�t	<\u��]���� ^_[�R�U�6����!Z�6�>f�:t	��!�����6�e��߀�@�6��f��f��f������f��f��6f���f��6���f��f�����fSf�f��6f���6f���f��f��f�� f�f���f3�6f���f��f[�S2۸�q�!s3ɋ�[�6�g׀ u��6��� ;6��� G6��� C2��6�0�6�2�%��'P�\ ��s� �\ 
�tJ���ҸJ� �/�uw�Àtr��SR�J�ӻ �/ZY�u]����:�t�ـ>Q� t�Ӵ6B�!=��t?�P���ʋ؃��� �JӉL�#�#��u ��u��t���NӸCq���)��!r�����SQR���tN��8��J r?f��f��f�f3���f�>ӡ8�:�<Ӌ@Ӌ>�B�D�F�H��F ZY[�3���� ��W�>'� uV�$����c��^�Cq����!��>'� uPR�ZX�_�VW����P�2�����< P���t#��3ҋ˻
 ����ыϋ������ыϋ�� [=
 rC3�����_^��%f��f��f��f��f+�f���f��f������3ҋ��SJ�LӊNӀ�r��3Ҁ��������������[úR��CN��������  �]� 3��� ��� �e��;"rW�>"�t;�|:u�| u� � ��ڀ>��v� ���x�`V��]�^������(붃>�� t� �i��S�����t�Gø �T��>�>��u��q� �5s�(�=  u�>�� u�u��4��#'� ����������� �؉�!s�tJ�NV��s�����e���]���� V��� 
�t���Vq���, ^�؉�F!s�t
���!� �� ���;�Gu�:8Du	�
 �����V�s�5 �%��x^rIV�%���>[��(5�%�6��<.u�| t%�|.t��e�i䡽��!^rV�%��b^s�2��^���= u� P����t�X= u����= u���e��������
sË6�V�%^�����@ 3ɺ����!s���׵�����^�ظ D�!�t�������B3ҋ��!������ B3��!��� ���3�.�>�� t�.���.�>���t1.�>�� t.)��.��� �.;��v.����c.���  �.)��?�!s�g����H3�P��X�=  u�}�u
+�I&���S� �@�![r;�u�z�I;�uû � D�!�u����&�� &��� �Q��jr&���&�>��v� �# �&�6����t�S��&�>���t�� � ��&���&���&��� ������ �L��s�tF��r?��� �{��W� ��/
�t#���/=��u���/��҉9�;�3۸��/_�����z� ��U ���3������G� �?�%մ�!�U ��!P�\ @<@u���A���X
�t����RӋ׾-չ �2����R�D2��\ �f���!Zr�1J� ����&J� ����������SQRVWU�Y3��!]_^ZY[ÿ݉�� �s� �t�� �~����S��? �[
�t4�3�!��<v��A��׺a��I���r	��r
�w���q���t��I�o�0�!P2䣗�X��2䣙׺ǂ�I�	r&�= u
�c �>�U �8&�G
�t0<$t�D ��&�G���
�t����	:t�? u���W�W_��ú��&I�=���
�>��<��|R�д�!Zô�!A������Ҁ6�t'2Ҿ!ԋ����AP���X�:���Ҫ���>�ס���!s�O��
�t
���`q���!������� �H�� �!ԉ�׺��HÊ\ ��@<@u�����!�P�:���ҪX�3P�$ԡ���!Xs�
�t�!ԋ��`q���!�!��2��~ΪO�u�� 2ۋ�3�<;uF� ����6gsbWVQ�YrU^V�<t0VW&�G��t"P��/����<tP��/��:�t�_^F��_^��؋�Y_PQ+ɬ��t�A��^�;�A�^_�<u�t �C<"u�� �tF�� u<;t
�P u��T ��O �Ѭ<t< t�<	t������ ������A�~ξ煬2�����t���`��� �(
��� u�9
�V2��D�P��W�e�<"t��
�u�_�e�|:\t2��,V��^
�t#WQ�b1Y_s2��2��+�Q3�I�Y�XP&�E�X^�D�t�����;�s�<"t�A�Iu�n��몬A����r!��2�+����"���4�+�󤺸��F�Z��F�� �/<�t/�D� ������!r�ۊ�	ۊ��>� � D�!�t��u�R �,��<v
<t� ���@ �؋J �6� 
�u���� �����R�3��Z3��ȴ�2�����3��þ�����2������!����Q��� �rV��٬�< u��}�:u�E� ^�r�"��ٸ l�B 3ɺ�!r�ظ D�!�u�>�!�`��E�@R���S�BG�ָ @� �![Zr�2����D�!S� 3۴>�!C��[�E�!�E�!�E�!�>�!���� ���P��� �/��hs�tr�m�#�=d rR=�wM�w���sJ�� �/<�t�I��I�w״f��!sP= u�Y3��!= u�L��(����#�Y3��!=A u�)��� ��� �w�������f��!�w׺9���D��� �V���s�t�<�~δ �.��5�~΃u�|:uV���:.�� �^�V��< u�^�Jr�S� ����~΋>޴`�!s�7����~�����z�މ�׺���/D���Ҁ> u�>� t��P˃>�t�v˸.��>1�3�/����&��&� &��&�
 &��&� &�i���u� P&�> u?&�r���t�.%�!&�t�¸/5�!C&�?Xu(&9Wu"��&�v���t�/%�!��ҴI�!X�L�!����ҴI�!J��&� X�1�!������� ���= t�6s׉6�����FúɅ�EC�̅�?C�� ���x �
�!�
�t�� ��4< tT<uֹJ��jQ�CYu����|:u9�P��/DD,@r+< w'� P�ط�D�!s� �DR3��!ZXr@��u�Ä�u������	��t����+��־텬2�ǻ � �����>�Ҁ v3QVW�Hۋt	�p�������+ƺޅ���� F؅�� 3��k _^Y���t�3�3�V�A<t<=u����À<u���^��t�X��^	������+ƳVP��%��=PAu�%��=THu�<=uX^��~�2�3����X^S��u3��� [��u��2������V���� ���� �u����^V�<t���3��^�>� t����ҿ���3�&�
P��Xu�|:u��$,A����&�
@&�>�&����&�>���\t<t<\u
&�>�&�����2���� O&�>�Î�ҎE3��< t��4謪
�u�����A��RPVP�1� � �  r	���o��+�X;�vY+�
�t4�>�Ҁ v,P�߅�1� � ���  r���@+�X+���ޅ�� @ PQ�YX�@Q���Y;�r
�؃�Q�cY^�� XZþ����V�1� �6 r���� ��&�< uN�z+��^þ��1� ����1� � r�� ��W ���&�E3�QVW�� tN�GG&;E�u'I���"�� G�>1� tS��&�E�� :�[�&:E�u��_^Yt�Q�h Y&�= u�����W�S �W�Q Y+���À<t���J t������j �<=u�þ ��Hۋt	��H
�t��<Î��&�E3�&�= t�� ���=�2����VP.����65�< tXP:rF:vF��3��3�@X^�<�r,�S����-���[�
<ar<zw, �PQS�d�[S���� �r�����Ȍ�;�w	˃� ;�s�J�!�����D��s����-[YX�P��H��&� �����X���Ҁ> t�С���!2��!��� �V���s�t+=��u]�
 �V�t^�u�|:u�mr�v� �/��7����b�V��	�R^s������u����!s��V�= t
= t�� ��.��O>��@ r:����!s6�1�= t(= t�� �!�~δ�!�N� �!r�����t�~���1��>��� ��� ��� :��� 9�Q��r7��ڀ>��w=��u3V�
^
�t���:q���9q�~�WV��< u�^_������t=��t����������� �;���>���t� �z����w�ru����!s�f�= t
= t�	 ��4��_=���������������ÿ�Њ�@<@u������&�Ұ:���&����!r��R������=Zì� t�N�< t�<=t�<,t�<;t�<	t�<
�S� �߬�< t�������.����< t������2��[æu��|� t���P�QW�������E� ��׺��<�E�r_YX���Һ��&�>C t��f��P�/8&��t</t<\X��(�  �>�� u	�\ ��rI��q�uB���u;�s&t�-r��u�(��= t = t= t�(�� �?�] � �2��� �6k�N�D��>G� u�{�td�΋�R;�t��]�tF���d�u��J���Z;�tb�΋�;�t�G� ��6�t�F�G���P�.8Du	8Dt�| Xt�D�<:t-�>G� u��u��� �a s:t�u r��u.�B �f�(��Ê����u�3ۆ|�8 st��L r���tڈ|� F�6�؜�q�u�\ �)�!��= u�\����!�V���	
��u����!������g	�^á��2��!Ó�3����:��u�� �F���<t�F���䈹 �u����I� ������̺����

�t���t)��ӵ �k��*����Ӻ����q���/�>�� u� �	�3�3۾�ӊG�:��u�k��t�C�� ��A������G�GG��>*� u�L�����% t&���t�>�u��  P��t	P���,:X��X�����t���
��<�u�<��T���u�9 t��������� ��6k����k���Ҁ>'�' t�)����P.�
��X���t���.��� .���.��ӿ��� ���t= }�Z�� ���'� �J���&�k��ҴI�!���&�&� ��ѿ��y�\ ���R ���gt0�J�/� �����с��J�/R��@� ��J�/Z�J�/�E� K����Ѽ��.�.�� starting
���t�&����3�!��r���������8�+�3�2Ҏ�ҋ6ދ��<"u����<u�R�����t��=  u��D�  ��<"u��t�<"u�����<u�Z������t�G��GAA��<"u��t�GA�<"u���<>ub8u�&�<�8�<<t<u�&��	 � W����2�Q��<t#<"u����
�u�	�t:��t<<t<>uN� Y�Z�����Y�M<<u%�����<>t<u��a�	 �YW�a��2�렊���|u9&�>C u&�&&�C��<t<|u�$��:+߃��t&8E�uO2��_�%G��tA��&�>C t'�H�A �!r4��3���҉>����6��E���<u�V�6މL�^���&�>C úՅ�e��� �6ދD��Hۋt	�P��X�Z �+Ƌ��{��r�~ ��G���<t���ρ� &�� Ë޾�Ӻ��� ��/< à]�
�t	VW�$!_^s
��VW�_^�PV3�3۬�Y#t<t�� u:��t
A<"u�� ��^Xái׋6k���� ��� �6k׋i׋��BQV[+�Yˋ��G� �1I����tIF�G������u���<?u���<*u���:��u�N2��F<u��6k׉i�þHۋL�t	� �� <ì��u<;u�N�RPSQVW.��Ҁ>C u�!�u� �Y _^Y[XZ�.�>a� t��a��.���@ � 3��!r�ذ��G� ��� =A t�t������.������.������>� u� ���H�>< tj.���" � � �!rP�ظ D�!�uh�B������!� ?� ����!r+;�u�>�����u@�B������!�4.��Ҹ B3ɋ��!�$= �u�[���3�.���" �� �!s�B��ذ��G� �P3���A
�u�+�X�s��SQVWUPR�Y�!Y[���=A t�Ë�]_^Y[������VWQ���3ɋ�W�< t�A�����_�i׉>k�Y_^�V�e�O�_�Q�2��Y^V�i��^�G+��Y�e��[�i�Y�[��S�_‿d�\t��d�:tƇe�\C�_��Y�e�Y�>]� uǇe�*.CC�_�Ǉe�* �_�[�W�>[⬪
�u�_�W�>Y���S�_�i�؀�\t�\C�[�[�VWQRU�6Y�V3��
�tE��^3ɋ�V�<.u��<*u��<?u��
�tA��^��� t�<.tG�>_�u
�>f�:u�4�]ZY_^Ã>_�u�|:u������>]� t�� u��.*�D �_����A��G����6Y��*�RQS
�u�RQS�(
�uX�|:�t��!A�:.�m�.�o�\ V�m��3۸�q3���!^rh�� @tb�.���;q.���Gq.���lq.���Cq�`W�|�<\\tGG�2������Ѱ\�|�  u�O2�&�P��3۸�q3���!X&�_�.��� ;.��� G.��� l.��� C2�[YZøNq.�>]� u�o � N3��!�u�r.�q�.�>]� t�u,��S�Oq.�q�.�>]� u�= � O3��!��u.�>]� t�u,��s.�>]� t�% �[ø\\9t�|:u9Dt2��R�״�!Z�S3�.�q��t��q�![þ%�.6��< u���.��҃> t�>��ð��Ң�R.��ҺE�A�!���A�!Z��D úX���������PR�f�����`1ZX�=A u��鶸����DWV��&�1� ���r���9�B�E� C�!r�� u�^_s	�. �E���E3ɴZ�!s눋ش>�!���Z�!s�v��ش>�!���6��>�u� ��3�&��6����<|t<|t�� �?� =�!s�1��ذ��G� &�>�3����<u���|��8t����<|t��������t���AA��<"uA����A<"u���<tEA<|t<|u�&�E�IW&�>�&�M�_N�6��AQ3ɸ <�!Ys�d��ذ��G� �?�A�W&�>�&�M�_N�6��Z��u��@��> �u����/��4��� ��s"�u<���p��r��^��/�� s�t��A�6C�D�QR�aZYs	�+�!
�u��Y�[��|/����t�9�3�3��w'=��t=  uN�#�,�!������������I/�=��t-=  u)�.I�J�6K�L�QR� 3��.'<�ZYu	�-�!
�u���񺎂�
/뿎�Ҁ&B�B�P.���3����t���I�!2��C
�t�.Xô,�!���.�3�.�5��.��.ø 8�\ �!��.��Hx.��t.���¶�.0.�6���~��..���  �
.�ރ�V�����D ^� �U��!V.�6ރ��L2�^3��=����x r�4�.��  �M�3ɺ���'.�
�2�����D � �
��!3������4�9�3�3��&���< t�<	t�Nù��2����W�_�ѿEWQ�Y_������PS�޻  �O [X�PSV�k�;M�t;�6����M�3��OۣKۢJ�V�6މ6Qۋ6H��� F�6��� ^��� ^[X�QRWV��t��  �����3�� �����3���׋6޹�+˿H���
��u�ދ���A������I�ޢ���Ҁ>" u!�>ދ�����A�W2�W�>ދM�_A�^�4� �6�3��&� ���(s	�t� ����� s��������6�^_ZY�SQRWV��n �>�@}W�����޸H�� �G  ��t�o�O�w�6މ76��+��w�6މw	���>��;>�}+��� ��>ޝ��
� �� ��^_ZY[�PSQW���t'�>��t&�޸H��# 	o	.
��� t������ޣޝ�_Y[X�P�ó��[ÓË�;�s6�H�����O�W�7�� uh� t&�Et����6#ӌ%Ӣ"�3���ø  �������t �uY�|:uS���,A��<wF�#Ӱ���@t	�[ r�r.���t�Q��}�t�E�?t�Ut��#�떸 �� ���뙸
 ����;�s�����PS�ӸH��(��7[X�W�w	���3��ӣӣӣ ��u|� ��� rt�w ��� ri�l �ӊ< t<	t<t<
uQ�Ӄ>�� t���t@�C�Ӄ>��r��t���t'�D�Ӄ>��r���t=d wl�A��L ��_�������3�Q�
 �<0r�<9w���� -0 F��Yì:��t</t
<-t<.t�À<0r��<9w����PS�ӸH��7�;w	w[X��Ӌ�;�s���VQR3ҹ
 �
�t,0<	w�������#��"��ZY^ø	 ���V�
�t�A�tFF���b�F��^�SQRVWU��J�t�� �o�޸ �!�5�Eމ>Gމ>I�������r֋��������tx�迊ы6HۋK�+΀>]� u[�:��r8dt�<\\tI���A������ �Ѐ�`�Kު��r
�I:K�tANV��� G�!��^�Kހ= t	�= tG���8uFI��� �J�u�K�+Hہѷ;�� ���u�s�3�� ��ыIދ6Gހ>]� uI�6E�&�|.u;&�|:u4Q&��&�D�E�KވEV&��� ��`����� G�!���N��^��Y+��6E�������tI��I��w��K����O�K�:E�t��N�3ɬ�
�u��ŝ]_^ZY[�P�/��t�\�K�XâK�X�SQRWV�����4�6Iމ6G��6E�2ɬ
�t:�t�r�t�����2ɪ���6Iފ�3���4�t!�K�
�u:E�t���
�u��4����]�� 
�t�����^_ZY[�QRWV������q��]��
�t�u� �$�re�L�  �N� �>]� t�����R ;L�~�L޿N޹ ��<u�w���ހ>]� t�u� ��s��Lހ>��t��ڣL��
����1'�  �^_ZY�WV�  ��� ��u?���e߁e�߾؈�ǧu�u� �#���܈�u�u� �������u�u� ��  �>��t=  t��ڸ ^_�PSQRWV�����Hۋ>Kۀ= t@��M�+�A�.���	��� � ��+ʃ�;�������>��u
O��?���� ������^_ZY[X��{sÀ>�� u�>]� t�e�Aq3ɾ �!ri�c�#�  �e�! �%���s�QR���r�i�A�>]� t�Aq3��!Zs��ۣ%���#�%���sɃ>#� u�%���>�� t����������R���Z�>��u��� ���`��� r���u�À&������ t������i��Ҋ������"� �3ru��!�2��!À� u�2��!�ںl����
�u���R�ڭ���&���7Z�B%�  ��P �.r
u�!����!À� t�N�T�!�غ|���+ r��Ҳ u��AÀ� t�%�&��ҊA���댿[��� ���r/�'��  2�_�� G:�u
QVW�_^Yt�@������s9���A�Hۋt	2����g t�N�<t#�\ t	�% u��V�<t�H t�^t���������7���< r+</t&<"t"<\t<:t<.t<[t<]t<+t
<,t<;t<=�:��< t<	�W� ��#��%����$_ÿ4�*�!�QR������Ρ{��W�%_� � �ZYÊ�2��#e�!À>�� ur�6Y�2Ҭ
�t&<*u��t�������<.uP��uK��tF���Հ�u	�>]� t5���u.����#�� ���x �
�!�
�t��g���r< t<��F�u����À>�� t(�5�i��׺��Q#����K#��!�T�r
< t<u�����3���ڹ 󪣼ڣ�٢�ң�ң�ң�ڣ�آ�Ң�آ�٢�آ�٢�آ�٢�Т�Ң'ۣ�Ң�ң�Ң�Ң�ڢ�ڣ�ף�أ�٣c֣%գ��H��٢�Ң�٢���� ��Hۋt	�u��+���3���׉6s��������ǀt�����u� ��w�� t��� t�� @�� t��� A(3) = 0
    Else
        A(1) = Left$(ValeurEntre, 1)
        A(2) = Mid$(ValeurEntre, 2, 1)
        A(3) = Len(ValeurEntre) - 2
    End If

    Rem Affiche les couleurs correspondantes
    frmR�sistor!opt1(A(1)).Value = True
    frmR�sistor!opt2(A(2)).Value = True
    frmR�sistor!opt3(A(3)).Value = True
    
    'Donne la tol�rance
    For X% = 0 To 5
        If optT(X%).Value Then
            frmR�sistor!opt4(X%).Value = -1
            Exit For
        End If
    Next
    
    frmCalcC.Hide
    Call frmR�sistor.cmdCalcV_Click
    Screen.MousePointer = 0
End Sub

Private Sub cmdOK_GotFocus()
    cmdOK.FontBold = True
End Sub

Private Sub cmdOK_LostFocus()
    cmdOK.FontBold = False
End Sub

Private Sub Form_Activate()
    txtValeur.SetFocus
End Sub

Private Sub optOhm_Click(Index As Integer, Value As Integer)
    'V�rifi si le valeur est bonne
    ValeurTexte = Val(txtValeur.Text) * 10 ^ Index
    If ValeurTexte < 1 Or ValeurTexte > 99000000 Then
        cmdOK.Enabled = False
    Else
        cmdOK.Enabled = True
    End If
End Sub

Private Sub pnlValeur_Click()
    txtValeur.SetFocus
End Sub

Private Sub txtValeur_Change()
    'Calcule la valeur du texte
    For i% = 0 To 6 Step 3
        If optOhm(i%) Then
            ValeurTexte = Val(txtValeur.Text) * 10 ^ i%
            Exit For
        End If
    Next
        
    'v�rifit si elle est juste
    If ValeurTexte < 1 Or ValeurTexte > 99000000 Then
        cmdOK.Enabled = False
    Else
        cmdOK.Enabled = True
    End If
End Sub

Private Sub txtValeur_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc(",")
            KeyAscii = Asc(".")
        Case Asc("0") To Asc("9"), Asc("."), 8, 13
            'Si caractere valide: n'ex�cute pas Else
        Case Else
            KeyAscii = 0
            Beep
    End Select
End Sub
