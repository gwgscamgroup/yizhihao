% 辅助Job2中的资管周报与长瑞周报填写
%% 参数
% 更新时需修改
begT=20190308;
endT=20190315;
useWsd=1;   % 是否要用wsd代替wss取数据
% 一般无需修改
location='D:\Job2\运维周报';
locHolder='D:\Job2\瑞益持有人份额';
inforPro={'长城国瑞证券长瑞1号定向资产管理计划','G-B-1','SG8892'; ...
          '长城国瑞证券瑞益2号定向资产管理计划','G-C-2','SG9617'; ...
          '长城国瑞证券瑞益1号集合资产管理计划','G-C-1','SQ4475'; ...
          '长城国瑞证券瑞盈1号定向资产管理计划','G-D-1','SQ0803'; ...
          '长城国瑞证券瑞益6号定向资产管理计划','G-D-6','SET595'};
inforType={'同业存单',0;'国债',1;'地方政府债',2;'央行票据',3;'政策性金融债',4;'政策银行债',4;'商业银行债',5; ...
'商业银行次级债',5;'保险公司债',5;'证券公司债',5;'证券公司短期融资券',5;'其他金融机构债',5; ...
'一般企业债',6;'集合企业债',6;'NPB',6;'一般公司债',7;'中期票据',8;'一般中期票据',8;'一般短期融资券',9; ...
'超短期融资债券',9;'PPN',10;'国际机构债',11;'政府支持机构债',12;'证监会主管ABS',13; ...
'银监会主管ABS',13;'交易商协会ABN',13;'项目收益票据PRN',13;'可转债',14;'可交换债',15; ...
'可分离转债存债',16;'非公开发行公司债',17;'私募债',17;'其他',18};
%% 
if useWsd
    fun='w.wsd';
else
    fun='w.wss';    %#ok<UNRCH>
end
locLtday=[location,num2str(begT),'\【3-2】CISP资管周报表',num2str(begT),'（投管）.xlsx'];
locToday=[location,num2str(endT)];
files=dir(locToday);
if exist([locToday,'\myHold.xlsx'],'file')
    [~,mySheets,~]=xlsfinfo([locToday,'\myHold.xlsx']);
    % 在确认完成第一步后，再确认代码有无重复
    numWrong=0;
    for n=1:length(files)
        filename=files(n).name;
        if strfind(filename,'证券投资基金估值表_');
            ind=strfind(filename,'_');
            namePro=filename(ind(1)+1:ind(2)-1);
            if ~ismember(namePro,mySheets)
                continue;
            end
            [~,indHold]=ismember(namePro,inforPro(:,1));
            [~,~,myHold]=xlsread([locToday,'\myHold.xlsx'],namePro);
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
        [~,~,data]=xlsread(locLtday,'3_5债券业务报表');   % 读取上一期的持仓数据作为筛选变动项的对照
        data([1:4,end],:)=[];
        data(:,1)=[];
        
        result=cell(0,15);
        for n=1:length(files)
            filename=files(n).name;
            if strfind(filename,'证券投资基金估值表_');
                ind=strfind(filename,'_');
                namePro=filename(ind(1)+1:ind(2)-1);
                if ~ismember(namePro,mySheets)
                    continue;
                end
                [~,indHold]=ismember(namePro,inforPro(:,1));
                [~,~,myHold]=xlsread([locToday,'\myHold.xlsx'],namePro);
                numHold=size(myHold,1);
                rltTmp=cell(numHold,15);
                rltTmp(:,1:3)=repmat(inforPro(indHold,[2,1,3]),numHold,1);  % 分别是编号、名称、产品代码
                eval(['typeHold=',fun,'(myHold(:,1),''windl2type'');']);
%                 typeHold=w.wss(myHold(:,1),'windl2type');
                [~,indType]=ismember(typeHold,inforType(:,1));
                indType(indType==0)=size(inforType,1);   % 所以要特别注意标号为18的，可能是搜索不到匹配的字段凑过来的
                rltTmp(:,4)=inforType(indType,2);   % 债券类型编码
                rltTmp(:,5)=myHold(:,2);    % 债券简称
                rltTmp(:,6)=myHold(:,1);    % 债券代码
                eval(['rltTmp(:,7)=',fun,'(myHold(:,1),''carrydate'');']);
                eval(['rltTmp(:,8)=',fun,'(myHold(:,1),''maturitydate'');']);
                eval(['rltTmp(:,9)=',fun,'(myHold(:,1),''amount'');']);
%                 rltTmp(:,7)=w.wss(myHold(:,1),'carrydate'); % 起息日
%                 rltTmp(:,8)=w.wss(myHold(:,1),'maturitydate');  % 到期日
%                 rltTmp(:,9)=w.wss(myHold(:,1),'amount');    % 债项评级
                rltTmp(cellfun(@isnumeric,rltTmp(:,9)),9)={'无'};
                eval(['rltTmp(:,10)=',fun,'(myHold(:,1),''latestissurercreditrating'');']);
%                 rltTmp(:,10)=w.wss(myHold(:,1),'latestissurercreditrating');    % 主体评级
                rltTmp(cellfun(@isnumeric,rltTmp(:,10)),10)={'无'};
                rltTmp(:,11:12)=myHold(:,4:5);  % 市值与成本
                eval(['rltTmp(:,13)=num2cell(',fun,'(myHold(:,1),''couponrate2'')/100);']);
%                 rltTmp(:,13)=num2cell(w.wss(myHold(:,1),'couponrate2')/100);  % 票面利率
                rltTmp(:,14)=repmat({'-'},numHold,1);   % 备注
                [~,indBR]=ismember(myHold(:,1),{'101778002.IB';'112494.SZ';'136388.SH'});
                rltTmp(indBR==1,14)={'违约'};
                rltTmp(indBR>=2,14)={'违约'};
                
                if strfind(namePro,'长瑞1号')
              %% 单独整理给委托人的周报
                    holdCR=cell(numHold,6);
                    holdCR(:,1:4)=myHold(:,1:4);
                    holdCR(:,5)=rltTmp(:,10);
                    if useWsd
                        holdCR(:,6)=w.wsd(myHold(:,1),'industry_sw','','','industryType=1');
                    else
                        holdCR(:,6)=w.wss(myHold(:,1),'industry_sw','industryType=1'); %#ok<UNRCH>
                    end                    
                    xlswrite([locToday,'\myHold.xlsx'],holdCR,'长瑞1号债券持仓');
                    % 评级整理
                    holdTmp=holdCR(~strcmp(holdCR(:,5),'无'),:);
                    valSum=sum(cell2mat(holdTmp(:,4)));
                    uni=unique(holdTmp(:,5));
                    prob=nan(length(uni),1);
                    for i=1:length(uni)
                        ind=find(strcmp(holdTmp(:,5),uni{i}));
                        prob(i)=sum(cell2mat(holdTmp(ind,4)))/valSum;
                    end
                    grade=cat(2,uni,num2cell(prob));
                    xlswrite([locToday,'\myHold.xlsx'],grade,'长瑞1号评级占比');
                    % 行业整理
                    holdTmp=holdCR(~cellfun(@isnumeric,holdCR(:,6)),:);
                    valSum=sum(cell2mat(holdTmp(:,4)));
                    uni=unique(holdTmp(:,6));
                    prob=nan(length(uni),1);
                    for i=1:length(uni)
                        ind=find(strcmp(holdTmp(:,6),uni{i}));
                        prob(i)=sum(cell2mat(holdTmp(ind,4)))/valSum;
                    end
                    idy=cat(2,uni,num2cell(prob));
                    xlswrite([locToday,'\myHold.xlsx'],idy,'长瑞1号行业占比');
                end
                indCd=find(indType==1);
                if ~isempty(indCd)
                    disp([namePro,'：',rltTmp(indCd,5)]);
                    rltTmp(indCd,:)=[];
                end
                result=cat(1,result,rltTmp);    % 资管周报债务持仓里不报同业存单
            end
        end
        % 筛选有有变动的项目
        file=unique(result(:,1));
        for n=1:length(file)
            ind0=find(strcmp(data(:,1),file{n}));
            ind1=find(strcmp(result(:,1),file{n}));
            for m=1:length(ind1)
                locb=find(strcmp(data(ind0,6),result{ind1(m),6}));
                if length(locb)>1
                    locb=locb(1);
                end
                if isempty(locb)
                    result{ind1(m),15}=0;
                else
                    if ~strcmp(data{ind0(locb),9},result{ind1(m),9})
                        result{ind1(m),15}=[result{ind1(m),15},'J'];
                    end
                    if ~strcmp(data{ind0(locb),10},result{ind1(m),10})
                        result{ind1(m),15}=[result{ind1(m),15},'K'];
                    end
                    if data{ind0(locb),11}~=result{ind1(m),11}
                        result{ind1(m),15}=[result{ind1(m),15},'L'];
                    end
                    if data{ind0(locb),12}~=result{ind1(m),12}
                        result{ind1(m),15}=[result{ind1(m),15},'M'];
                    end
                    if data{ind0(locb),13}~=result{ind1(m),13}
                        result{ind1(m),15}=[result{ind1(m),15},'N'];
                    end
                    if ~strcmp(data{ind0(locb),14},result{ind1(m),14})
                        result{ind1(m),15}=[result{ind1(m),15},'O'];
                    end
                end
            end
        end
        % 瑞益1号持有人信息
        fileHolder=dir(locHolder);
        [~,indMax]=max(arrayfun(@(x) x.datenum,fileHolder(3:end)));
        [~,~,tmp]=xlsread([locHolder,'\',fileHolder(indMax+2).name]);
        nameHolder=tmp(2:end-1,2);
        numHolder=cell2mat(tmp(2:end-1,9));
        [numSort,indSort]=sort(numHolder,'descend');
        probHolder=cat(2,nameHolder(indSort),num2cell([numSort,numSort/sum(numSort)]));
        xlswrite([locToday,'\myHold.xlsx'],probHolder,'持有人份额');
        
        % 输出所有结果
        xlswrite([locToday,'\myHold.xlsx'],result,'result');
    end
else
    % 第一步，先将几个产品的债券持仓集中取出，手动补充wind代码
    % 手动完成这步后，再次运行此文件
    for n=1:length(files)
        filename=files(n).name;
        if strfind(filename,'证券投资基金估值表_');
            ind=strfind(filename,'_');
            namePro=filename(ind(1)+1:ind(2)-1);
            [~,~,data]=xlsread([locToday,'\',filename]);
            indDebt=cellfun(@(x) ischar(x) && length(x)>4 && (strcmp(x(1:4),'1103') || strcmp(x(1:4),'1104')),data(:,1) );
            indValid=cellfun(@(x) isnumeric(x) && ~isnan(x),data(:,4));
            indChs=indDebt & indValid;
            myHold=data(indChs,[2,3,8,5,1]);% 分别提取简称、数量、市值、成本、科目代码
            
%             indS=find(cellfun(@(x) strcmp(x,'债券投资'),data(:,2)));
%             indE=find(cellfun(@(x) length(x)>2 && strcmp(x(1:2),'12'),data(:,1)),1);
%             indValid=find(cellfun(@(x) isnumeric(x) && ~isnan(x),data(:,4)));
%             indChs=find(indValid>indS & indValid<indE);
%             myHold=data(indValid(indChs),[2,3,8,5,1]);  % 分别提取简称、数量、市值、成本、科目代码
            if strfind(namePro,'长瑞1号')
%                 myHold=cat(1,myHold,{'16凯迪03',500000,50000000,50000000,0;'16亿阳04',70000,7000000,7000000,0;'17永泰能源MTN001',700000,70000000,70000000,0});
                myHold=cat(1,myHold,{'16凯迪03',500000,50000000,50000000,0});
            end
            if ~isempty(myHold)
                xlswrite([locToday,'\myHold.xlsx'],myHold,namePro);
            end
        end
    end
    disp('已集中取出债券简称，手动补充wind代码后需再次运行本文件');
end