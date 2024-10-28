% 设置输入和输出文件夹路径
inputFolder = 'E:\2023国自科多模态\VA研究\02-分离sheet';
outputFolder = 'E:\2023国自科多模态\VA研究\06-VA共合并';

% 创建输出文件夹，如果不存在
if ~exist(outputFolder, 'dir')
    mkdir(outputFolder);
end

% 获取输入文件夹下所有 Excel 文件
files = dir(fullfile(inputFolder, '*.xlsx'));

% 初始化存储所有文件处理后数据的表格
combinedData = [];

% 遍历每一个 Excel 文件
for i = 1:length(files)
    % 读取 Excel 文件
    fileName = files(i).name;
    filePath = fullfile(inputFolder, fileName);
    
    % 读取 Excel 文件中的数据
    data = readtable(filePath);
    
    % 初始化用于保存纵向合并的数据
    combinedCols1 = [];
    combinedCols2 = [];

    % 按照每两列依次纵向合并
    for col = 1:2:width(data) % 每次处理两列，步长为2
        if col+1 <= width(data)
            % 提取成对的列
            tempCols1 = data{:, col};   % 第一个列对的第1列
            tempCols2 = data{:, col+1}; % 第一个列对的第2列
        else
            % 如果最后剩下单独一列，提取这一列，第二列用空值填充
            tempCols1 = data{:, col};
            tempCols2 = NaN(size(tempCols1)); % 填充空值
        end
        
        % 纵向合并这两列数据
        combinedCols1 = [combinedCols1; tempCols1];
        combinedCols2 = [combinedCols2; tempCols2];
    end
    
    % 将处理后的数据纵向拼接
    combinedData = [combinedData; table(combinedCols1, combinedCols2)];
end

% 生成新的 Excel 文件路径
outputFile = fullfile(outputFolder, 'combine_VA.xlsx');

% 将合并后的数据输出到新的 Excel 文件中
writetable(combinedData, outputFile, 'WriteVariableNames', false);

disp('所有文件处理完毕，已保存至 combine_VA.xlsx 文件中。');
