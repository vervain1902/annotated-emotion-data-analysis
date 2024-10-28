% 设置文件夹路径
inputFolder = 'E:\2023国自科多模态\VA研究\02-分离sheet';
outputFolder = 'E:\2023国自科多模态\VA研究\03-合并列';

% 获取文件夹下所有 Excel 文件
files = dir(fullfile(inputFolder, '*.xlsx'));

% 遍历每一个 Excel 文件
for i = 1:length(files)
    % 读取 Excel 文件
    fileName = files(i).name;
    filePath = fullfile(inputFolder, fileName);
    
    % 读取 Excel 文件中的数据
    [~, sheets] = xlsfinfo(filePath); % 获取工作表信息
    for j = 1:length(sheets)
        data = readtable(filePath, 'Sheet', sheets{j});
        
        % 初始化奇数列和偶数列的存储
        oddColumns = [];
        evenColumns = [];
        
        % 提取奇数列和偶数列
        for col = 1:width(data)
            if mod(col, 2) == 1
                oddColumns = [oddColumns; data{:, col}];  % 奇数列
            else
                evenColumns = [evenColumns; data{:, col}]; % 偶数列
            end
        end
        
        % 构建新的表格
        oddTable = array2table(oddColumns);
        evenTable = array2table(evenColumns);
        
        % 生成新的文件名
        [~, baseFileName, ~] = fileparts(fileName);
        oddFileName = fullfile(outputFolder, [baseFileName, '_odd.xlsx']);
        evenFileName = fullfile(outputFolder, [baseFileName, '_even.xlsx']);
        
        % 保存奇数列数据
        writetable(oddTable, oddFileName, 'WriteVariableNames', false);
        
        % 保存偶数列数据
        writetable(evenTable, evenFileName, 'WriteVariableNames', false);
    end
end

disp('所有文件处理完毕并已保存。');
