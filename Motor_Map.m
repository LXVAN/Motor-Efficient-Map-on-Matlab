function Motor_Map()
    % 电机效率MAP工具 - 最终物理边界版
    
    % --- 1. 创建主界面窗口 ---
    fig = uifigure('Name', '电机效率数据分析系统 v2.0', 'Position', [400 300 600 650]);

    % 初始化全局存储变量
    appData = struct('table', [], 'fileName', '');

    % --- 2. 界面组件布局 ---
    % 文件路径区
    uilabel(fig, 'Text', '1. 数据源路径:', 'Position', [20 610 100 22], 'FontWeight', 'bold');
    efPath = uieditfield(fig, 'text', 'Position', [20 580 450 30], 'Placeholder', '请输入或选择文件路径...');
    uibutton(fig, 'Text', '浏览...', 'Position', [480 580 100 30], 'ButtonPushedFcn', @(btn,event) browseFile());

    % 数据列选择区
    uilabel(fig, 'Text', '2. 设置数据映射列:', 'Position', [20 540 300 22], 'FontWeight', 'bold');
    uilabel(fig, 'Text', '转速 (Speed):', 'Position', [20 510 100 22]);
    ddSpeed = uidropdown(fig, 'Position', [130 510 450 22]);
    uilabel(fig, 'Text', '转矩 (Torque):', 'Position', [20 475 100 22]);
    ddTorque = uidropdown(fig, 'Position', [130 475 450 22]);
    uilabel(fig, 'Text', '效率 (Efficiency):', 'Position', [20 440 100 22]);
    ddEff = uidropdown(fig, 'Position', [130 440 450 22]);

    % 数据范围显示区 (回答你的第2点要求)
    uilabel(fig, 'Text', '3. 数据极值统计 (自动提取):', 'Position', [20 400 200 22], 'FontWeight', 'bold');
    lblRangeS = uilabel(fig, 'Text', '转速范围: --', 'Position', [30 375 550 22], 'FontColor', [0.4 0.4 0.4]);
    lblRangeT = uilabel(fig, 'Text', '扭矩范围: --', 'Position', [30 355 550 22], 'FontColor', [0.4 0.4 0.4]);
    lblRangeE = uilabel(fig, 'Text', '效率范围: --', 'Position', [30 335 550 22], 'FontColor', [0.4 0.4 0.4]);

    % 样式与输出
    uilabel(fig, 'Text', '4. 图形样式与输出设置:', 'Position', [20 290 200 22], 'FontWeight', 'bold');
    uilabel(fig, 'Text', '配色方案:', 'Position', [20 260 80 22]);
    ddColor = uidropdown(fig, 'Position', [100 260 150 22], 'Items', {'jet', 'parula', 'turbo', 'hot', 'cool'}, 'Value', 'jet');
    uilabel(fig, 'Text', '导出格式:', 'Position', [300 260 80 22]);
    ddFormat = uidropdown(fig, 'Position', [380 260 150 22], 'Items', {'.png', '.jpg', '.pdf', '.tif'}, 'Value', '.png');

    % 按钮
    uibutton(fig, 'Text', '预览生成 MAP 图', 'Position', [20 120 560 45], 'FontSize', 14, 'FontWeight', 'bold', 'BackgroundColor', [0.9 0.95 1], 'ButtonPushedFcn', @(btn,event) processPlot(false));
    uibutton(fig, 'Text', '保存图片', 'Position', [20 60 560 45], 'FontSize', 14, 'BackgroundColor', [0.9 1 0.9], 'ButtonPushedFcn', @(btn,event) processPlot(true));

    % --- 3. 回调函数 ---
    function browseFile()
        [file, pth] = uigetfile({'*.xlsx;*.xls;*.csv', '数据文件'}, '选择文件');
        if isequal(file, 0), return; end
        fullPath = fullfile(pth, file);
        efPath.Value = fullPath; 
        updateDataLogic(fullPath);
    end

    efPath.ValueChangedFcn = @(ef, event) updateDataLogic(ef.Value);

    function updateDataLogic(targetPath)
        if isfile(targetPath)
            try
                opts = detectImportOptions(targetPath, 'VariableNamingRule', 'preserve');
                appData.table = readtable(targetPath, opts);
                [~, appData.fileName, ~] = fileparts(targetPath);
                
                cols = appData.table.Properties.VariableNames;
                ddSpeed.Items = cols; ddTorque.Items = cols; ddEff.Items = cols;
                
                % 默认选中前三列并更新统计信息
                if length(cols) >= 3
                    ddSpeed.Value = cols{2}; ddTorque.Value = cols{3}; ddEff.Value = cols{12};
                end
                refreshStats();
            catch ME
                uialert(fig, ['读取失败：', ME.message], '错误');
            end
        end
    end

    % 统计数据范围
    function refreshStats()
        try
            sVal = str2double(regexprep(string(appData.table.(ddSpeed.Value)), '[^\d\.\-]', ''));
            tVal = str2double(regexprep(string(appData.table.(ddTorque.Value)), '[^\d\.\-]', ''));
            eVal = str2double(regexprep(string(appData.table.(ddEff.Value)), '[^\d\.\-]', '')) / 10;
            lblRangeS.Text = sprintf('转速范围: Min = %.0f rpm, Max = %.0f rpm', min(sVal,[],'omitnan'), max(sVal,[],'omitnan'));
            lblRangeT.Text = sprintf('扭矩范围: Min = %.0f Nm, Max = %.0f Nm', min(tVal,[],'omitnan'), max(tVal,[],'omitnan'));
            lblRangeE.Text = sprintf('效率范围: Min = %.1f%%, Max = %.1f%%', min(eVal,[],'omitnan'), max(eVal,[],'omitnan'));
        catch
        end
    end

    % 切换下拉框时自动更新统计
    ddSpeed.ValueChangedFcn = @(src, event) refreshStats();
    ddTorque.ValueChangedFcn = @(src, event) refreshStats();
    ddEff.ValueChangedFcn = @(src, event) refreshStats();
    

    function processPlot(isSave)
        % 檢查數據是否存在
        if isempty(appData.table)
            uialert(fig, '請先加載有效的数据文件！', '提示'); 
            return; 
        end
        
        try
            % 1. 獲取當前選中的列名
            sCol = ddSpeed.Value;
            tCol = ddTorque.Value;
            eCol = ddEff.Value;
            
            % 2. 數據提取與清洗 (處理特殊字符並轉為數值)
            S_raw = str2double(regexprep(string(appData.table.(sCol)), '[^\d\.\-]', ''));
            T_raw = str2double(regexprep(string(appData.table.(tCol)), '[^\d\.\-]', ''));
            E_raw = str2double(regexprep(string(appData.table.(eCol)), '[^\d\.\-]', '')) / 10;
            
            % 剔除 NaN 行
            mask = ~isnan(S_raw) & ~isnan(T_raw) & ~isnan(E_raw);
            S = S_raw(mask); T = T_raw(mask); E = E_raw(mask);
            
            if isempty(S), uialert(fig, '數據轉換後為空，請檢查列選擇！', '錯誤'); return; end

            % 3. 提取外特性包絡線並進行【平滑擬合】
            % 以 10rpm 為步長聚合，消除測試時的轉速微小抖動
            [unqS, ~, idx] = unique(round(S/10)*10); 
            maxT_raw = accumarray(idx, T, [], @max);
            
            % 使用 6 階多項式擬合包絡線（消除狗牙感）
            p_env = polyfit(unqS, maxT_raw, 6);

            % 4. 插值生成網格 (嚴格限制在數據範圍內)
            sMin = min(S); sMax = max(S);
            tMax_data = max(T);
            
            s_grid = linspace(sMin, sMax, 300);
            t_grid = linspace(0, tMax_data * 1.05, 300); % 縱軸稍微留白以便顯示曲線
            [X, Y] = meshgrid(s_grid, t_grid);
            Z = griddata(S, T, E, X, Y, 'v4');

            % 5. 【物理邊界裁剪】抹除擬合線以上的所有插值區域
            limitT_smooth = polyval(p_env, X); 
            Z(Y > limitT_smooth) = NaN;

            % 6. 開始繪圖
            hFig = figure('Color', 'w', 'Name', appData.fileName);
            if isSave, hFig.Visible = 'off'; end
            
            % 調整繪圖區邊距，減少外部留白
            set(gca, 'LooseInset', get(gca, 'TightInset'));
            
            % 繪製效率雲圖 (0-100級別)
            levels = 0:1:100; 
            [C, hC] = contourf(X, Y, Z, levels, 'LineStyle', 'none'); 
            hold on;
            colormap(ddColor.Value);
            
            % 繪製平滑後的黑色外特性曲線
            s_env_line = linspace(sMin, sMax, 300);
            t_env_line = polyval(p_env, s_env_line);
            plot(s_env_line, t_env_line, 'k-', 'LineWidth', 2.5);
            
            % 7. 【核心優化】Colorbar 與 0-100% 刻度
            cb = colorbar;
            cb.Label.String = '效率 (%)';
            cb.Label.FontWeight = 'bold';
            
            % 設置顏色映射範圍：從數據最小值到 100%
            caxis([min(E) 100]); 
            % 手動強制設置 0-100 刻度
            cb.Ticks = 0:10:100; 
            drawnow;
            cb.TickLabels = strcat(cellstr(num2str(cb.Ticks', '%.0f')), '%');
            
            % 等高線標籤
            hT = clabel(C, hC, 'FontSize', 8);
            if ~isempty(hT)
                for i = 1:length(hT)
                    txtStr = get(hT(i), 'String');
                    if ~contains(txtStr, '%'), set(hT(i), 'String', [txtStr, '%']); end
                end
            end
            
            % 8. 【核心優化】座標軸範圍與最大值顯示 (解決右側空白問題)
            xlabel([sCol, ' (rpm)'], 'FontWeight', 'bold');
            ylabel([tCol, ' (Nm)'], 'FontWeight', 'bold');
            title(appData.fileName, 'Interpreter', 'none');
            grid on;
            
            % 嚴格鎖定 X 軸到 sMax，不再乘 1.05
            axis([sMin sMax 0 tMax_data * 1.05]);
            
            % 手動指定刻度，確保最大轉速剛好在最右側邊界
            xticks_val = linspace(sMin, sMax, 6);
            xticks(xticks_val);
            xtickformat('%.0f'); % 格式化為整數
            
            % 9. 圖片導出邏輯
            if isSave
                [outF, outP] = uiputfile(['*', ddFormat.Value], '保存圖片', [appData.fileName, ddFormat.Value]);
                if ~isequal(outF, 0)
                    exportgraphics(hFig, fullfile(outP, outF), 'Resolution', 300);
                    uialert(fig, '圖片已成功導出！', '成功', 'Icon', 'success');
                end
                delete(hFig);
            end
            
        catch ME
            uialert(fig, ['生成 MAP 圖時出錯：', ME.message], '錯誤');
            if exist('hFig','var') && ishandle(hFig), delete(hFig); end
        end
    end
end