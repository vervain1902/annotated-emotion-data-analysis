% 设置文件夹路径
inputFolder = 'E:\2023国自科多模态\VA研究\03-合并列';
outputFolder = 'E:\2023国自科多模态\VA研究\04-odd-V';

% 创建输出文件夹，如果不存在
if ~exist(outputFolder, 'dir')
    mkdir(outputFolder);
end

% 获取文件夹下所有 Excel 文件
files = dir(fullfile(inputFolder, '*odd*.xlsx')); % 只匹配名称含有'odd'的文件

% 初始化保存最终统计结果的表格
finalResults = [];

% 遍历每一个符合条件的 Excel 文件
for i = 1:length(files)
    % 读取 Excel 文件
    fileName = files(i).name;
    filePath = fullfile(inputFolder, fileName);
    
    % 读取 Excel 文件中的第一列数据
    data = readtable(filePath, 'ReadVariableNames', false);
    firstColumn = data{:, 1}; % 获取第一列数据
    
    % 将第一列数据从 cell 类型转换为适当的类型
    if iscell(firstColumn)
        if iscellstr(firstColumn)
            firstColumn = string(firstColumn);  % 如果是字符串，转换为字符串数组
        else
            firstColumn = cell2mat(firstColumn);  % 如果是数值，转换为数值数组
        end
    end
    
    % 初始化频数统计
    uniqueValues = [];
    frequencies = [];
    
    % 遍历第一列数据，统计相同值的连续出现频率
    count = 1;
    for j = 2:length(firstColumn)
        if firstColumn(j) == firstColumn(j-1)
            count = count + 1; % 连续相同数值计数增加
        else
            % 保存前一个数值和对应频数
            uniqueValues = [uniqueValues; firstColumn(j-1)];
            frequencies = [frequencies; count];
            count = 1; % 重置计数器
        end
    end
    % 处理最后一个数值
    uniqueValues = [uniqueValues; firstColumn(end)];
    frequencies = [frequencies; count];
    
    % 将当前文件的统计结果添加到最终结果
    finalResults = [finalResults; table(uniqueValues, frequencies)];
end

% 生成新的 Excel 文件名
outputFile = fullfile(outputFolder, 'V-frequency.xlsx');

% 保存统计结果到新的 Excel 文件中
writetable(finalResults, outputFile, 'WriteVariableNames', false);

disp('频数统计已完成并保存至 "V-frequency.xlsx" 文件中。');
