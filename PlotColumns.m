%% Plot the data in excel worksheet and link the brush of subplots.
% Parse the inputs sent from excel and call the generator below to create figure.
%
% Varaibles sent from excel:
% data1, data2: data of subplots, numerical array (without headers) or
%   cell of char and number (with headers)
% optionName1, optionName2: names of plot setting options, 1st columns of the
%   setting tables in .xslm file, m by 1 char cell
% optionVal1, optionVal2: values of plot setting options, 2nd columns of the
%   setting tables in .xslm file, m by 1 char cell

if ~exist("data1", "var")
    return
end

if exist("optionName1", "var") && iscell(optionName1)...
    && exist("optionVal1", "var") && iscell(optionVal1)
    validOptions = false(size(optionName1));
    for j = 1:length(optionVal1(:))
        if isnumeric(optionVal1{j})
            validOptions(j) = ~(isnan(optionVal1{j}) || optionVal1{j} == 0);
        else
            validOptions(j) = true;
        end
    end
    optName1 = strcat(optionName1(validOptions), {'1'});
    options1 = [optName1 optionVal1(validOptions)]';
else
    options1 = {};
end

if exist("data2", "var")
    if exist("optionVal2", "var") && iscell(optionVal2) ...
        && exist("optionName2", "var") && iscell(optionName2)
        validOptions = false(size(optionName2));
        for j = 1:length(optionVal2(:))
            if isnumeric(optionVal2{j})
                validOptions(j) = ~(isnan(optionVal2{j}) || optionVal2{j} == 0);
            else
                validOptions(j) = true;
            end
        end
            optName2 = strcat(optionName2(validOptions), {'2'});
            options2 = [optName2 optionVal2(validOptions)]';
            generatePlot(data1, data2, options1{:}, options2{:});
    else
        generatePlot(data1, data2, options1{:});
    end
else
    generatePlot(data1, options1{:});
end

function generatePlot(data1, data2, opt1, opt2)
% generatePlot is the implementation of plotting.
arguments
    data1
    data2 = []
    opt1.title1 = ''        % for subplot 1
    opt1.xlabel1 = ''
    opt1.ylabel1 = ''
    opt1.axis1 = 'on'
    opt1.lineSpec1 = '-'
    opt1.grid1 = 'off'
    opt2.title2 = ''        % for subplot 2
    opt2.xlabel2 = ''
    opt2.ylabel2 = ''
    opt2.axis2 = 'on'
    opt2.lineSpec2 = '-'
    opt2.grid2 = 'off'
end

set(groot,'defaultFigureColor','w')
set(groot, 'defaultAxesFontSize', 12)
set(groot, 'defaultLineLineWidth', 1)
figure()

tiledlayout(1, 1 + ~isempty(data2), TileSpacing="tight", Padding="compact")
nexttile
options = struct2cell(opt1);
p1 = draw(data1, options{:});

if ~isempty(data2)
    nexttile
    options = struct2cell(opt2);
    p2 = draw(data2, options{:});

    % Setup brush sync
    if all(size(cell2mat({p1.XData}')) == size(cell2mat({p2.XData}')))
        b = brush(gcf);
        b.ActionPostCallback = {@onBrushOneToOne, p1, p2};
    elseif size(cell2mat({p1.XData}'), 2) == size(cell2mat({p2.XData}'), 2)
        if size(cell2mat({p1.XData}'), 1) == 1
            b = brush(gcf);
            b.ActionPostCallback = {@onBrushAllToOne, p2, p1};
        elseif size(cell2mat({p2.XData}'), 1) == 1
            b = brush(gcf);
            b.ActionPostCallback = {@onBrushAllToOne, p1, p2};
        end
    end
end

end

function [p] = draw(p_data, p_title, p_xlabel, p_ylabel, p_axis, p_lineSpec, p_grid)
% Plot the data
%
% Parameters:
% p_data: columns of data to be plotted, 1st column is x-axis
% p_title: title of the plot
% p_xlabel: label of x-axis
% p_ylabel: label of y-axis
% p_axis: {equal | tight | padded | on | off} (can choose multiple options)
% p_lineSpec: marker, linestyle, colour
% p_grid: {on | off | minor}
%
% Note: The order of parameters here MUST match the argument block in function
% `generatePlot`.

arguments
    p_data
    p_title = ''
    p_xlabel = ''
    p_ylabel = ''
    p_axis = 'on'
    p_lineSpec = '-'
    p_grid = 'off'
end

% Generate headers
default_header = cellstr("data " + string(1:size(p_data, 2)-1));
if iscell(p_data)                       % p_data contains string
    p_header = p_data(1, 2:end);
    p_data = cell2mat(p_data(2:end, :));
    p_header(~cellfun(@istext, p_header)) = default_header(~cellfun(@istext, p_header));
else
    p_header = default_header;
end

% plot
hold on
for j=2:size(p_data, 2)
    plot(p_data(:, 1), p_data(:, j), p_lineSpec)
end
hold off
p = get(gca, 'Children');

% show legend if more than one line
if size(p_data, 2) > 2
    legend(p_header)
end

% additional options
title(p_title)
xlabel(p_xlabel)
ylabel(p_ylabel)
grid(p_grid)
if (isa(p_axis, "char") || isa(p_axis, "string"))
    p_axis = cellstr(split(p_axis, " "));
    axis(p_axis{:})
end

end

function output = istext(S)
% istext Determine whether input is text
% iscellstr(S) returns 1 if S is text and 0 otherwise.
output = isstring(S) || ischar(S)|| iscellstr(S);
end

% callback function
% Sync the 'BrushData' when data of whichever axis are brushed.
% equal number of lines for the axes, n vs. n
function onBrushOneToOne(~, eventdata, p1, p2)
if p1(1).Parent == eventdata.Axes
    set(p2, {'BrushData'}, {p1.BrushData}');
else
    set(p1, {'BrushData'}, {p2.BrushData}');
end
end

% callback function
% Sync the 'BrushData' when data of whichever axis are brushed.
% n vs. 1
function onBrushAllToOne(~, eventdata, p1, p2)
if p1(1).Parent == eventdata.Axes
    bData = uint8(any(cell2mat({p1.BrushData}')));
    set(p2, 'BrushData', bData);
else
    set(p1, {'BrushData'}, {p2.BrushData}');
end
end