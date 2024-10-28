% 设置输入和输出文件夹路径
inputFolder = 'E:\2023国自科多模态\VA研究\06-VA共合并';
outputFolder = 'E:\2023国自科多模态\VA研究\07-VA_Frequency';

% 创建输出文件夹，如果不存在
if ~exist(outputFolder, 'dir')
    mkdir(outputFolder);
end

% 获取输入文件夹下所有 Excel 文件
files = dir(fullfile(inputFolder, '*.xlsx'));

% 初始化保存统计结果的表格
finalResults = [];

% 遍历每一个 Excel 文件
for i = 1:length(files)
    % 读取 Excel 文件
    fileName = files(i).name;
    filePath = fullfile(inputFolder, fileName);
    
    % 读取 Excel 文件中的第一列和第二列数据
    data = readtable(filePath);
    
    % 提取第一列和第二列数据
    col1 = data{:, 1}; % 第一列数据
    col2 = data{:, 2}; % 第二列数据
    
    % 初始化统计结果的变量
    uniqueValues1 = [];
    uniqueValues2 = [];
    frequencies = [];
    
    % 初始化统计连续出现相同数值的计数器
    count = 1;
    
    % 遍历数据，统计连续出现相同数值的频数
    for j = 2:length(col1)
        if col1(j) == col1(j-1) && col2(j) == col2(j-1)
            count = count + 1; % 相同数据对，计数器增加
        else
            % 保存前一个数据对和对应的频数
            uniqueValues1 = [uniqueValues1; col1(j-1)];
            uniqueValues2 = [uniqueValues2; col2(j-1)];
            frequencies = [frequencies; count];
            
            % 重置计数器
            count = 1;
        end
    end
    % 处理最后一对数据
    uniqueValues1 = [uniqueValues1; col1(end)];
    uniqueValues2 = [uniqueValues2; col2(end)];
    frequencies = [frequencies; count];
    
    % 将当前文件的统计结果添加到最终结果
    finalResults = [finalResults; table(uniqueValues1, uniqueValues2, frequencies)];
end

% 生成新的 Excel 文件路径
outputFile = fullfile(outputFolder, 'VA_Frequency.xlsx');

% 将统计结果输出到新的 Excel 文件中
writetable(finalResults, outputFile, 'WriteVariableNames', false);

disp('频数统计已完成并保存至 VA_Frequency.xlsx 文件中。');
