clc
clear

%读取3D坐标
[BIM3D, No_a] = xlsread('Data.xlsx',1);
BIM3D(:,4)=1;

%投影
K = xlsread('Data.xlsx',2);
R_t = xlsread('Data.xlsx',3);
BIM2D = (K * R_t * BIM3D.').';
BIM2D = BIM2D./(BIM2D(:,3)*[1,1,1]);

%读取检测结果
[Detect_result, No_b] = xlsread('Data.xlsx',4);

[Rows_a, Columns_a] = size(BIM3D); 
[Rows_b, Columns_b] = size(Detect_result); 
BIM_element_id = No_a(1,1);
Envelope_box = [];
Final_result = {};
for n = 1:Rows_b   %对于每一条检测结果数据
   Component_type = No_b(n,2); %检测到的构件类型
   Top_view = Detect_result(n,1); %是否为俯视图
   Layer = No_b(n,4); %拍摄楼层   
   Position = Detect_result(n,3:6); %检测结果位置
   Position_width = Position(2)-Position(1);
   Position_height = Position(4)-Position(3);
   D = [];
   D_index = {};
   k = 0;

   for m = 1:Rows_a
       if ~strcmp(No_a(m,2), Component_type) %判断构件类型
           continue;
       else
           if Top_view == 1 && ~strcmp(No_a(m,3), Layer) %判断俯视图
               continue;
           else
               if strcmp(No_a(m,1),BIM_element_id) %判断构件楼层
                   [Envelope_box_rows, Envelope_box_Columns] = size(Envelope_box);                   
                   Envelope_box(Envelope_box_rows+1, 1:2) = BIM2D(m, 1:2); %包络框增加一个点
               else
                   %此时，已经获得一个包络框
                   W_max = max(Envelope_box(:,1));
                   W_min = min(Envelope_box(:,1));
                   W_value= W_max - W_min;
                   H_max = max(Envelope_box(:,2));
                   H_min = min(Envelope_box(:,2));
                   H_walue = H_max - H_min;
                   if W_max > Position(1) && W_min < Position(1) && H_max>Position(3) && H_min<Position(4)  %有重叠
                       if W_value<2*Position_width && H_walue < 2*Position_height && 2*W_value>Position_width && 2*H_walue>Position_height %尺度一致
                           %计算
                           d = zeros(Envelope_box_rows+1,1);
                           k = k+1;
                           D(k,1) = 0;
                           for i = 1:Envelope_box_rows+1   %对于每一个点
                               %计算其对应的距离
                               dd = zeros(4,1);
                               dd(1,1) = abs(Envelope_box(i,1)-Position(1));
                               dd(2,1) = abs(Envelope_box(i,1)-Position(2));
                               dd(3,1) = abs(Envelope_box(i,1)-Position(3));
                               dd(4,1) = abs(Envelope_box(i,1)-Position(4));
                               d(i,1)=min(dd);
                               D = D + d(i,1) * d(i,1);
                           end
                           D(k,1) = D(k,1) / (Envelope_box_rows+1);
                           D_index(k,1) = BIM_element_id;
                       end
                   end
                  BIM_element_id = No_a(m,1);
               end
           end
       end
   end
   if k~=0
       [minvalue, index] = min(D(:,1));
       Final_result = cat(1, Final_result, [No_b(n,1),D_index(index,1)])
   else
       disp('空')  
   end
end
xlswrite('Data_result.xlsx',Final_result)
