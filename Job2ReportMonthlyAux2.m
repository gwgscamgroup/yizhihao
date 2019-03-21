% 辅助Job2中压力测试的表格填写
%% 参数
endT='2019/2/28';
openNext={'长城国瑞证券瑞益1号集合资产管理计划','2019/3/7'};
          %'长城国瑞证券长瑞1号定向资产管理计划','2018/9/17'};    % 仅对180天未到期且内有开放的

numBeta=243;
numHsl=20;
degreeD=[-0.43,-0.55,-0.67];
codeIndex={'000001.SH';'399005.SZ';'399006.SZ'};
indIndex=[1;1;2;3];
location='D:\Job2\压力测试';
location=[location,datestr(endT,'yyyymm')];      
      
myTitle={'产品名称(1)','产品净值(2)','产品资产总值(3)','产品单位净值(4)','产品份额(5)', ...
'股票资产规模(6)','上证主板规模(7)','上证主板Beta(8)','深证主板规模(9)', ...
'深圳主板Beta(10)','中小板规模(11)','中小板Beta(12)','创业板规模(13)', ...
'创业板Beta(14)','公募基金资产规模(15)','债券资产规模(16)','利率债资产规模(17)', ...
'利率债久期(18)','信用债资产规模AAA(19)','信用债资产规模AA+(20)', ...
'信用债资产规模AA(21)','信用债久期AAA(22)','信用债久期AA+(23)','信用债久期AA(24)', ...
'现金规模(25)','可变现资产规模轻度(26)','可变现资产规模中度(27)','可变现资产规模重度(28)'};

%% 前提准备
files=dir(location);
numPro=0;
namePro=cell(0,1);  % 产品名称
dataPro=cell(0,1);  % 产品估值表数据
debtPro=cell(0,1); % 产品债券持仓代码、简称、持仓规模
numWrong=0;
listWrong=cell(numWrong,2);
% 主要先对债券的证券代码进行提取和辨别
he=actxserver('Excel.Application');
hw=he.Workbooks.Add;
he.Visible=1;
hs=hw.Worksheets;
sheetItem=hs.Item(1);
for n=1:length(files)
    filename=files(n).name;
    if strfind(filename,'证券投资基金估值表')        
        numPro=numPro+1;
        disp(['第',num2str(numPro),'只产品准备：']);
        ind=strfind(filename,'_');
        namePro{numPro,1}=filename(ind(1)+1:ind(2)-1);
        [~,~,dataPro{numPro,1}]=xlsread([location,'\',filename]);
        indDebt=find(cellfun(@(x) ischar(x) && length(x)>=14 && strcmp(x(1:4),'1103') ,dataPro{numPro}(:,1) ));
        if isempty(indDebt)
            continue;
        end
        debtPro{numPro,1}=cell(0,3);
        indSH=find(cellfun(@(x) str2double(x(1:6))<=110320 ,dataPro{numPro}(indDebt,1) ));
        if ~isempty(indSH)
            codeSH=cellfun(@(x) [x(9:end),'.SH'],dataPro{numPro}(indDebt(indSH),1) ,'UniformOutput',false);
            nameSH=dataPro{numPro}(indDebt(indSH),2);
            mvSH=dataPro{numPro}(indDebt(indSH),8);
            debtPro{numPro}=cat(1,debtPro{numPro},[codeSH,nameSH,mvSH]);
        end
        indSZ=find(cellfun(@(x) str2double(x(1:6))>110320 && str2double(x(1:6))<=110340 ,dataPro{numPro}(indDebt,1) ));
        if ~isempty(indSZ)
            codeSZ=cellfun(@(x) [x(9:end),'.SZ'],dataPro{numPro}(indDebt(indSZ),1) ,'UniformOutput',false);
            nameSZ=dataPro{numPro}(indDebt(indSZ),2);
            mvSZ=dataPro{numPro}(indDebt(indSZ),8);
            debtPro{numPro}=cat(1,debtPro{numPro},[codeSZ,nameSZ,mvSZ]);
        end
        % 利用excel服务器读取证券代码，错误有两种情况
        % 1、后缀与所属场所不一致；2、直接读不出（因名称有误或有空格）
        indIB=find(cellfun(@(x) str2double(x(1:6))>110340 ,dataPro{numPro}(indDebt,1) ));
        if ~isempty(indIB)
            nameIB=strrep(dataPro{numPro}(indDebt(indIB),2),' ','');
            mvIB=dataPro{numPro}(indDebt(indIB),8);
            codeIB=cell(length(indIB),1);
            for m=1:length(indIB)
                sheetItem.Range('A1').Value=['=to_windcode("',nameIB{m},'")'];
                codeIB{m}=sheetItem.Range('A1').Value;
                if isnumeric(codeIB{m}) % || ~strcmp(codeIB{m}(end-1:end),'IB')   % 后面的条件是为了精确匹配银行间，但如果不提取价格其实不需要
                    numWrong=numWrong+1;
                    listWrong(numWrong,:)={namePro{numPro},nameIB{m}};
                end
            end
            debtPro{numPro}=cat(1,debtPro{numPro},[codeIB,nameIB,mvIB]);
        end
        xlswrite([location,'\myHold.xlsx'],debtPro{numPro},namePro{numPro});
    end
end
disp(['有',num2str(numWrong),'个债券代码有问题！']);

clear he hw hs sheetItem;
save Job2ReportMonthlyAux2.mat;

%% 
% load Job2ReportMonthlyAux2.mat;
if ~numWrong
    ts=actxserver('TSExpert.CoExec');
    w=windmatlab;
    begT=w.tdaysoffset(-numBeta,endT);
    tmp=w.wsd(codeIndex,'close',begT,endT);
    priceIndex=tmp(:,indIndex);
    retIndex=priceIndex(2:end,:)./priceIndex(1:end-1,:)-1;
    
    result=cell(0,length(myTitle));
    for n=1:length(files)
        filename=files(n).name;
        if strfind(filename,'证券投资基金估值表')            
            rlt=cell(1,length(myTitle));
            ind=strfind(filename,'_');
            rlt{1}=filename(ind(1)+1:ind(2)-1); % 产品名称
            [~,locb]=ismember(rlt{1},namePro);
            myData=dataPro{locb};
            myDebt=debtPro{locb};
            disp([rlt{1},'提取计算：']);
            
            ind2=find(cellfun(@(x) ischar(x) && ~isempty(strfind(x,'基金资产净值')),myData(:,1)));
            rlt{2}=myData{ind2,8};   % 产品净值
            ind3=find(cellfun(@(x) ischar(x) && ~isempty(strfind(x,'资产类合计')),myData(:,1)));
            rlt{3}=myData{ind3,8};    % 资产总值
            ind4=find(cellfun(@(x) ischar(x) && ~isempty(strfind(x,'基金单位净值')),myData(:,1)));
            rlt{4}=myData{ind4,2};    % 基金单位净值
            ind5=find(cellfun(@(x) ischar(x) && ~isempty(strfind(x,'实收资本金额')),myData(:,1)));
            rlt{5}=myData{ind5,8};    % 基金产品份额
            ind6=find(cellfun(@(x) ischar(x) && strcmp(x,'1102'),myData(:,1)));
            if ~isempty(ind6)
                rlt{6}=myData{ind6,8};    % 股票资产规模
            end
            ind7=find(cellfun(@(x) ischar(x) && length(x)==14 && strcmp(x(1:9),'110201016') ,myData(:,1)));
            if ~isempty(ind7)
                code=cellfun(@(x) [x(end-5:end),'.SH'],myData(ind7,1),'UniformOutput',false);
                price=w.wsd(code,'close',begT,endT,'PriceAdj=F');
                ret=price(2:end,:)./price(1:end-1,:)-1;
                amt=w.wsd(code,'amt',begT,endT);
                bet=ones(length(code),1);
                indValid=find(prod(amt(2:end,:))>0 & ~isnan(prod(ret)) );
                numValid=length(indValid);
                if numValid>0
                    retValid=ret(:,indValid);
                    bet(indValid)=sum((retValid-repmat(mean(retValid),numBeta,1)).*repmat(retIndex(:,1)-mean(retIndex(:,1)),1,numValid))/(numBeta-1)/var(retIndex(:,1));
                end
                mv=cell2mat(myData(ind7,8));
                rlt{7}=sum(mv);     % 上证主板规模
                rlt{8}=bet'*mv/sum(mv); % 上证主板Beta
            end
            ind9=find(cellfun(@(x) ischar(x) && length(x)==14 && strcmp(x(1:11),'11023101000') ,myData(:,1)));
            if ~isempty(ind9)
                code=cellfun(@(x) [x(end-5:end),'.SZ'],myData(ind9,1),'UniformOutput',false);
                price=w.wsd(code,'close',begT,endT,'PriceAdj=F');
                ret=price(2:end,:)./price(1:end-1,:)-1;
                amt=w.wsd(code,'amt',begT,endT);
                bet=ones(length(code),1);
                indValid=find(prod(amt(2:end,:))>0 & ~isnan(prod(ret)) );
                numValid=length(indValid);
                if numValid>0
                    retValid=ret(:,indValid);
                    bet(indValid)=sum((retValid-repmat(mean(retValid),numBeta,1)).*repmat(retIndex(:,2)-mean(retIndex(:,2)),1,numValid))/(numBeta-1)/var(retIndex(:,2));
                end
                mv=cell2mat(myData(ind9,8));
                rlt{9}=sum(mv);     % 深证主板规模
                rlt{10}=bet'*mv/sum(mv); % 深证主板Beta
            end
            ind11=find(cellfun(@(x) ischar(x) && length(x)==14 && strcmp(x(1:11),'11023101002') ,myData(:,1)));
            if ~isempty(ind11)
                code=cellfun(@(x) [x(end-5:end),'.SZ'],myData(ind11,1),'UniformOutput',false);
                price=w.wsd(code,'close',begT,endT,'PriceAdj=F');
                ret=price(2:end,:)./price(1:end-1,:)-1;
                amt=w.wsd(code,'amt',begT,endT);
                bet=ones(length(code),1);
                indValid=find(prod(amt(2:end,:))>0 & ~isnan(prod(ret)) );
                numValid=length(indValid);
                if numValid>0
                    retValid=ret(:,indValid);
                    bet(indValid)=sum((retValid-repmat(mean(retValid),numBeta,1)).*repmat(retIndex(:,3)-mean(retIndex(:,3)),1,numValid))/(numBeta-1)/var(retIndex(:,3));
                end
                mv=cell2mat(myData(ind11,8));
                rlt{11}=sum(mv);     % 深证中小板规模
                rlt{12}=bet'*mv/sum(mv); % 深证中小板Beta
            end
            ind13=find(cellfun(@(x) ischar(x) && length(x)==14 && strcmp(x(1:9),'110241013'),myData(:,1)));
            if ~isempty(ind13)
                code=cellfun(@(x) [x(end-5:end),'.SZ'],myData(ind13,1),'UniformOutput',false);
                price=w.wsd(code,'close',begT,endT,'PriceAdj=F');
                ret=price(2:end,:)./price(1:end-1,:)-1;
                amt=w.wsd(code,'amt',begT,endT);
                bet=ones(length(code),1);
                indValid=find(prod(amt(2:end,:))>0 & ~isnan(prod(ret)) );
                numValid=length(indValid);
                if numValid>0
                    retValid=ret(:,indValid);
                    bet(indValid)=sum((retValid-repmat(mean(retValid),numBeta,1)).*repmat(retIndex(:,4)-mean(retIndex(:,4)),1,numValid))/(numBeta-1)/var(retIndex(:,4));
                end
                mv=cell2mat(myData(ind13,8));
                rlt{13}=sum(mv);     % 创业板规模
                rlt{14}=bet'*mv/sum(mv); % 创业板Beta
            end
            ind15=find(cellfun(@(x) ischar(x) && length(x)>=14 && strcmp(x(1:4),'1105') && ~isnan(str2double(x(9:end))), myData(:,1) ));
            if ~isempty(ind15)
                rlt{15}=sum(cell2mat(myData(ind15,8)));       % 公募基金资产规模      
            end
            
            ind16=find(cellfun(@(x) ischar(x) && strcmp(x,'1103'),myData(:,1)));
            if ~isempty(ind16)
                rlt{16}=myData{ind16,8};    % 债券资产规模
                
%                 baseRate=w.wss(myDebt(:,1),'baserate');      %% 浮息债基准利率
                baseRate=w.wsd(myDebt(:,1),'baserate','ED-0D',endT);      %% 浮息债基准利率
                isChange=logical(cellfun(@ischar,baseRate));
%                 dur=w.wss(myDebt(:,1),'modifiedduration',['tradeDate=',endT]);
                dur=w.wsd(myDebt(:,1),'modifiedduration','ED-0D',endT);
                dur(isChange)=0.5;  %% 对浮息债久期取值为0.5
%                 rateDebt=w.wss(myDebt(:,1),'latestissurercreditrating2,rate_ratebond',['tradeDate=',endT],'ratingAgency=101','type=1');   %% 主体、债项评级
                tmp1=w.wsd(myDebt(:,1),'latestissurercreditrating2','ED-0D',endT,'ratingAgency=101','type=1');
                tmp2=w.wsd(myDebt(:,1),'rate_ratebond','ED-0D',endT,'ratingAgency=101','type=1');
                rateDebt=cat(2,tmp1,tmp2);
                
%                 issuer=w.wss(myDebt(:,1),'issuerupdated'); %% 发行主体
                issuer=w.wsd(myDebt(:,1),'issuerupdated','ED-0D',endT); %% 发行主体
                indRat=find(cellfun(@(x) ismember(x,{'国家开发银行';'中国农业发展银行';'中国进出口银行';'中华人民共和国财政部'}) || strcmp(x(end-3:end),'人民政府') ,issuer)); %% 利率债下标
                if ~isempty(indRat)
                    mv=cell2mat(myDebt(indRat,3));
                    rlt{17}=sum(mv);    % 利率债资产规模
                    rlt{18}=dur(indRat)'*mv/sum(mv);    % 利率债久期
                end
                indCre=setdiff(1:size(myDebt,1),indRat)';   %% 信用债下标
                if ~isempty(indCre)
                    rate=rateDebt(indCre,:);
                    ind=cellfun(@isnumeric,rate(:,2));
                    rate(ind,2)=rate(ind,1);
                    rate(:,1)=[];
                    
                    ind3A=find(strcmp(rate,'AAA'));
                    if ~isempty(ind3A)
                        mv3A=cell2mat(myDebt(indCre(ind3A),3));
                        rlt{19}=sum(mv3A);      % AAA信用债规模
                        rlt{22}=dur(indCre(ind3A))'*mv3A/sum(mv3A); % AAA信用债久期
                    end
                    ind2AP=find(strcmp(rate,'AA+'));
                    if ~isempty(ind2AP)
                        mv2AP=cell2mat(myDebt(indCre(ind2AP),3));
                        rlt{20}=sum(mv2AP); % AA+信用债规模
                        rlt{23}=dur(indCre(ind2AP))'*mv2AP/sum(mv2AP);  % AA+信用债久期
                    end
                    ind2A=setdiff(1:length(indCre),union(ind3A,ind2AP))';
                    if ~isempty(ind2A)
                        mv2A=cell2mat(myDebt(indCre(ind2A),3));
                        rlt{21}=sum(mv2A);  % AA及其他信用债规模
                        rlt{24}=dur(indCre(ind2A))'*mv2A/sum(mv2A); % AA及其他信用债久期
                    end
                end
            end
            ind25=cellfun(@(x) ischar(x) && ismember(x,{'1002','1021','1031'}),myData(:,1));
            rlt{25}=sum(cell2mat(myData(ind25,8)));   % 现金规模
            
            [~,locb]=ismember(rlt{1},openNext(:,1));
            if locb         %% 仅针对在180日内未到期且有开放的计算可变现资产
                dateNext=openNext{locb,2};  %% 下一开放日
                numD=double(w.tdayscount(endT,dateNext)-1);
                % 可变现资产分为4类
                mv1=rlt{25};
                mv2=0;
                if ~isempty(ind16)
%                     dateMat=w.wss(myDebt(:,1),'maturitydate');
                    dateMat=w.wsd(myDebt(:,1),'maturitydate','ED-0D',endT);
                    indEarly=find(datenum(dateMat)<datenum(dateNext));
                    if ~isempty(indEarly)
                        mv1=mv1+sum(cell2mat(myDebt(indEarly,3)));  % 1、开放前到期金融资产（包括现金）
                    end
                    
                    indLate=setdiff(1:size(myDebt,1),indEarly)';
                    if ~isempty(indLate)
%                         typeHold=w.wss(myDebt(indLate,1),'windl2type');
                        typeHold=w.wsd(myDebt(indLate,1),'windl2type','ED-0D',endT);
                        ind1=find(ismember(typeHold,{'国债','政策银行债','央行票据'}));
                        rate=rateDebt(indLate,1);
                        ind2=find(cellfun(@(x) ~isempty(strfind(x,'短期融资')) ,typeHold) & strcmp(rate,'AAA') );
                        mv2=mv2+sum(cell2mat(myDebt(indLate(union(ind1,ind2)),3))); % 2、开放前未到期但短期可变现的债券
                    end
                end
                
                mv3=zeros(1,3);
                code=cell(0,1);
                indSH=ind7;
                if ~isempty(indSH)
                    code=cat(1,code,cellfun(@(x) ['SH',x(9:14)] ,myData(indSH,1),'UniformOutput',false));
                end
                indSZ=[ind9;ind11;ind13];
                if ~isempty(indSZ)
                    code=cat(1,code,cellfun(@(x) ['SZ',x(9:14)], myData(indSZ,1),'UniformOutput',false));
                end
                indCode=[indSH;indSZ];
                if ~isempty(indCode)
                    mvHold=cell2mat(myData(indCode,8));
                    hsl=mean(cell2mat(ts.RemoteCallFunc('pdFreeTurn4',{code,numHsl,str2double(datestr(endT,'yyyymmdd'))})))'/100;
                    mvFree=cell2mat(ts.RemoteCallFunc('pdStockMarketValue2',{code,str2double(datestr(endT,'yyyymmdd'))}))*1e4;
                    mv3=mv3+sum(min(repmat(mvHold,1,3),(hsl.*mvFree)*(1+degreeD)*0.05*numD ));% 3、股票可变现部分 
                end
                
                mv4=0;
                ind=find(cellfun(@(x) ischar(x) && length(x)>8 && strcmp(x(1:8),'11090301') ,myData(:,1)));
                if ~isempty(ind)
                    mv4=mv4+sum(cell2mat(myData(ind,8)));   % 4、公募资金、现金类理财产品
                end
                if ~isempty(rlt{15})
                    mv4=mv4+rlt{15};
                end
                rlt{26}=mv1+mv2+mv3(1)+mv4; % 轻度可变现资产规模
                rlt{27}=mv1+mv2+mv3(2)+mv4; % 中度可变现资产规模
                rlt{28}=mv1+mv2+mv3(3)+mv4; % 重度可变现资产规模
            end
            result=cat(1,result,rlt);
        end
    end
    
    result=cat(1,myTitle,result);
    xlswrite([location,'\myHold.xlsx'],result,'汇总');
end