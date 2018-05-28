%   ���û�ѡ���EXCEL�ļ����ݽ��кϲ�

% �û�ѡȡ�ļ��У���ȡ�ļ�����Ӧ��Ϣ
filePath = strcat(uigetdir([],'��ѡ�����Excel�ļ����ļ���'), '\');
fileList = dir(filePath);
nFiles = length(fileList);

% �������ļ������ƴ洢��һ��������
fileNameList = repmat({[]}, [1 nFiles]);
for i = 1:nFiles
    fileNameList{i} = fileList(i).name;
end

% �ų����з�Excel��ʽ���ļ�
invalidFileIndex = [];
for i = 1:nFiles
    if length(fileNameList{i}) > 4
        if ~strcmp(fileNameList{i}(end-3:end), '.xls') && ~strcmp(fileNameList{i}(end-4:end), '.xlsx')
            invalidFileIndex = [invalidFileIndex i];
        end
    else
        invalidFileIndex = [invalidFileIndex i];
    end
end
fileNameList(invalidFileIndex) = [];

% �������һ�����ϵ�Excel�ļ�����кϲ��������������������ʾ
combinedContent = {};
if length(fileNameList) > 1    
    for i = 1:length(fileNameList)
        [~, ~, currentContent] = xlsread(strcat(filePath, fileNameList{i}));
        % �ڶ����Ժ���ļ������һ�е�����
        if i > 1
            currentContent(1,:) = [];
        end
        
        % ���ļ�����ͬǰ���Ѿ���ȡ�����ݽ��кϲ�
        combinedContent = [combinedContent; currentContent];
        
    end    
    if xlswrite(strcat(filePath, 'all.xls'), combinedContent)
        msgbox('Excel�ļ��ϲ��ɹ�', '��ʾ'); 
    else        
        msgbox('�ļ�д��ʧ��', '����'); 
    end
    
elseif ~isempty(fileNameList) > 0
    msgbox('�ļ���ֻ��һ��Excel��ʽ���ļ���������кϲ�', '����');    
else
    msgbox('�ļ�����û��Excel��ʽ���ļ�', '����');
end

