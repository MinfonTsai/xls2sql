#include <stdio.h>
#include <string>
#include <afxsock.h>
#include "mysql.h" 

using namespace std;
class ToMysql 
{
public:
//�ܼ�
MYSQL mysql; 
/*
�c�y��ƩM�}�c���
*/
//VspdCToMySQL();
//~VspdCToMySQL();
/*
�D�n���\��G
��l�Ƹ�Ʈw
�s����Ʈw
�]�m�r����
�J�f�ѼơG
host �GMYSQL���A��IP
port:��Ʈw��
Db�G��Ʈw�W��
user�G��Ʈw�ϥΪ�
passwd�G��Ʈw�ϥΪ̪��K�X
charset�G�Ʊ�ϥΪ��r����
Msg:��^�������A�]�A���~����
�X�f�ѼơG
int �G0��ܦ��\�F1��ܥ���
*/
int ConnMySQL(char *host,char * port,char * Db,char * user,char* passwd,char * charset,char * Msg);
/*
�D�n���\��G
�d�߸��
�J�f�ѼơG
SQL�G�d�ߪ�SQL�y�y
Cnum:�d�ߪ��C��
Msg:��^�������A�]�A���~����
�X�f�ѼơG
string �ǳƩ�m��^����ơA�h���O���h��0x06�j�},�h������0x05�j�}
�p�G ��^�����ס� 0�A�d��ܻR���G
*/
string SelectData(char * SQL,int Cnum ,char * Msg);
/*
�D�n�\��G
���J���
�J�f�Ѽ�
SQL�G�d�ߪ�SQL�y�y
Msg:��^�������A�]�A���~����
�X�f�ѼơG
int �G0��ܦ��\�F1��ܥ���
*/
int InsertData(char * SQL,char * Msg);
/*
�D�n�\��G
�ק���
�J�f�Ѽ�
SQL�G�d�ߪ�SQL�y�y
Msg:��^�������A�]�A���~����
�X�f�ѼơG
int �G0��ܦ��\�F1��ܥ���
*/
int UpdateData(char * SQL,char * Msg);
/*
�D�n�\��G
�R�����
�J�f�Ѽ�
SQL�G�d�ߪ�SQL�y�y
Msg:��^�������A�]�A���~����
�X�f�ѼơG
int �G0��ܦ��\�F1��ܥ���
*/
int DeleteData(char * SQL,char * Msg);
/*
�D�n�\��G
������Ʈw�s��
*/
void CloseMySQLConn();
};
