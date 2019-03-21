% 辅助Job2中三个产品持仓月报填写
%% 参数
location='D:\Job2\持仓月报201902';
%%
files=dir(location);
if exist([location,'\myHold.xlsx'],'file')
    % 在确认完成第一步后，再确认代码有无重复
    numWrong=0;
    for n=1:length(files)
        filename=files(n).name;
        if strfind(filename,'证券投资基金估值表_');
            ind=strfind(filename,'_');
            namePro=filename(ind(1)+1:ind(2)-1);
%             if strfind(namePro,'瑞益1号')
%                 continue;
%             end
            namePro=namePro(7:10);
            [~,~,myHold]=xlsread([location,'\myHold.xlsx'],namePro);
            % 先处理有重复wind代码（这只是暂时的处理方式，以后再改善）
            [~,~,ic]=unique(myHold(:,1));
            tmp=tabulate(ic);
            indRep=find(ismember(ic,tmp(tmp(:,2)>1,1)));
            for i=1:length(indRep)
                ind=indRep(i);
                if (strcmp(myHold{ind,1}(end-1:end),'IB') && isnumeric(myHold{ind,6})) || ...
                   (~strcmp(myHold{ind,1}(end-1:end),'IB') && ischar(myHold{ind,6}))
                    disp([namePro,'第',num2str(ind),'只债券的代码有问题！']);
                    numWrong=numWrong+1;
                end
            end
            
        end
    end
    disp(['共',num2str(numWrong),'只债券的代码有问题']);
    % 
    if ~numWrong
        w=windmatlab;       
        for n=1:length(files)
            filename=files(n).name;
            if strfind(filename,'证券投资基金估值表_');
                ind=strfind(filename,'_');
                namePro=filename(ind(1)+1:ind(2)-1);
%                 if strfind(namePro,'瑞益1号')
%                     continue;
%                 end
                namePro=namePro(7:10);
                [~,~,myHold]=xlsread([location,'\myHold.xlsx'],namePro);
                numHold=size(myHold,1);
                holdInfor=cell(numHold,6);
                holdInfor(:,1:4)=myHold(:,1:4);
%                 holdInfor(:,5)=w.wss(myHold(:,1),'latestissurercreditrating');  % 主体评级
%                 holdInfor(:,6)=w.wss(myHold(:,1),'amount');     % 债项评级
%                 holdInfor(:,7)=w.wss(myHold(:,1),'industry_sw','industryType=1');
                holdInfor(:,5)=w.wsd(myHold(:,1),'latestissurercreditrating');  % 主体评级
                holdInfor(:,6)=w.wsd(myHold(:,1),'amount');     % 债项评级
                holdInfor(:,7)=w.wsd(myHold(:,1),'industry_sw','','','industryType=1');
                xlswrite([location,'\myHold.xlsx'],holdInfor,namePro);
                % 评级整理（无债项评级时，用其主体评级代替）
                grade=cell(numHold,1);
                indValid=~cellfun(@isnumeric,holdInfor(:,6));
                grade(indValid)=holdInfor(indValid,6);
                grade(~indValid)=holdInfor(~indValid,5); 
                holdInfor(:,6)=grade;
                holdTmp=holdInfor(~cellfun(@isnumeric,grade),[1:4,6:7]);
                valSum=sum(cell2mat(holdTmp(:,4)));
                uni=unique(holdTmp(:,5));
                prob=nan(length(uni),1);
                for i=1:length(uni)
                    ind=find(strcmp(holdTmp(:,5),uni{i}));
                    prob(i)=sum(cell2mat(holdTmp(ind,4)))/valSum;
                end
                grade=cat(2,uni,num2cell(prob));
                xlswrite([location,'\myHold.xlsx'],grade,[namePro,'评级']);
                % 行业整理
                holdTmp=holdInfor(~cellfun(@isnumeric,holdInfor(:,7)),:);
                valSum=sum(cell2mat(holdTmp(:,4)));
                uni=unique(holdTmp(:,7));
                prob=nan(length(uni),1);
                for i=1:length(uni)
                    ind=find(strcmp(holdTmp(:,7),uni{i}));
                    prob(i)=sum(cell2mat(holdTmp(ind,4)))/valSum;
                end
                idy=cat(2,uni,num2cell(prob));
                xlswrite([location,'\myHold.xlsx'],idy,[namePro,'行业']);
            end
        end
        
    end
else
    % 第一步，先将几个产品的债券持仓集中取出，手动补充wind代码
    % 手动完成这步后，再次运行此文件
    for n=1:length(files)
        filename=files(n).name;
        if strfind(filename,'证券投资基金估值表_')
            ind=strfind(filename,'_');
            namePro=filename(ind(1)+1:ind(2)-1);
%             if strfind(namePro,'瑞益1号')
%                 continue;
%             end
            namePro=namePro(7:10);
            [~,~,data]=xlsread([location,'\',filename]);
            indDebt=cellfun(@(x) ischar(x) && length(x)>4 && (strcmp(x(1:4),'1103') || strcmp(x(1:4),'1104')),data(:,1) );
            indValid=cellfun(@(x) isnumeric(x) && ~isnan(x),data(:,4));
            indChs=indDebt & indValid;
            myHold=data(indChs,[2,3,8,5,1]);% 分别提取简称、数量、市值、成本、科目代码
            if strfind(namePro,'长瑞1号')
                myHold=cat(1,myHold,{'16凯迪03',500000,50000000,50000000,0});
            end            
            xlswrite([location,'\myHold.xlsx'],myHold,namePro);
        end
    end
    disp('已集中取出债券简称，手动补充wind代码后需再次运行本文件');
end