%   将用户选择的EXCEL文件内容进行合并

% 用户选取文件夹，读取文件夹相应信息
filePath = strcat(uigetdir([],'请选择包含Excel文件的文件夹'), '\');
fileList = dir(filePath);
nFiles = length(fileList);

% 将所有文件的名称存储在一个数组中
fileNameList = repmat({[]}, [1 nFiles]);
for i = 1:nFiles
    fileNameList{i} = fileList(i).name;
end

% 排除所有非Excel格式的文件
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

% 如果包含一个以上的Excel文件则进行合并操作，否则给出错误提示
combinedContent = {};
if length(fileNameList) > 1    
    for i = 1:length(fileNameList)
        [~, ~, currentContent] = xlsread(strcat(filePath, fileNameList{i}));
        % 第二个以后的文件清除第一行的内容
        if i > 1
            currentContent(1,:) = [];
        end
        
        % 将文件内容同前面已经读取的内容进行合并
        combinedContent = [combinedContent; currentContent];
        
    end    
    if xlswrite(strcat(filePath, 'all.xls'), combinedContent)
        msgbox('Excel文件合并成功', '提示'); 
    else        
        msgbox('文件写入失败', '错误'); 
    end
    
elseif ~isempty(fileNameList) > 0
    msgbox('文件夹只有一个Excel格式的文件，无需进行合并', '错误');    
else
    msgbox('文件夹中没有Excel格式的文件', '错误');
end

