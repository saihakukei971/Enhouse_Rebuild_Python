@echo off
cd /d %~dp0

:: Python �X�N���v�g�̎��s - �L���gID�擾��CSV�o��
echo [INFO] Enhouse_01_�L���gID�擾��CSV�o��.py �����s��...
python Enhouse_01_�L���gID�擾��CSV�o��.py
echo [INFO] Enhouse_01_�L���gID�擾��CSV�o��.py ����

:: Python �X�N���v�g�̎��s - CSV�f�[�^���X�v���b�h�V�[�g�ɃA�b�v���[�h
echo [INFO] Enhouse_02_CSV�f�[�^���X�v���b�h�V�[�g�ɃA�b�v���[�h.py �����s��...
python Enhouse_02_CSV�f�[�^���X�v���b�h�V�[�g�ɃA�b�v���[�h.py
echo [INFO] Enhouse_02_CSV�f�[�^���X�v���b�h�V�[�g�ɃA�b�v���[�h.py ����

:: 3���i180�b�j�̑ҋ@�����i�A�b�v���[�h����ُ̈�f�[�^�ǉ���҂j
echo [INFO] 3���ҋ@���i�ُ�f�[�^���ǉ�����Ȃ����m�F�j
timeout /t 180
echo [INFO] 3���ҋ@����

:: Python �X�N���v�g�̎��s - �ُ�l�폜����
echo [INFO] Enhouse_03_�ُ�l�폜.py �����s��...
python Enhouse_03_�ُ�l�폜.py
echo [INFO] Enhouse_03_�ُ�l�폜.py ����

:: 3���i180�b�j�̑ҋ@�����i�폜��ُ̈�f�[�^�ǉ���҂j
echo [INFO] 3���ҋ@���i�ُ�f�[�^���ǉ�����Ȃ����Ċm�F�j
timeout /t 180
echo [INFO] 3���ҋ@����

:: Python �X�N���v�g�̎��s - �s�Ɗ֐��̎����ǉ�����
echo [INFO] Enhouse_04_�s�Ɗ֐��̎����ǉ�.py �����s��...
python Enhouse_04_�s�Ɗ֐��̎����ǉ�.py
echo [INFO] Enhouse_04_�s�Ɗ֐��̎����ǉ�.py ����

exit
