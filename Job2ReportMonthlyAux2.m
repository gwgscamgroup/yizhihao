% ����Job2��ѹ�����Եı����д
%% ����
endT='2019/2/28';
openNext={'���ǹ���֤ȯ����1�ż����ʲ�����ƻ�','2019/3/7'};
          %'���ǹ���֤ȯ����1�Ŷ����ʲ�����ƻ�','2018/9/17'};    % ����180��δ���������п��ŵ�

numBeta=243;
numHsl=20;
degreeD=[-0.43,-0.55,-0.67];
codeIndex={'000001.SH';'399005.SZ';'399006.SZ'};
indIndex=[1;1;2;3];
location='D:\Job2\ѹ������';
location=[location,datestr(endT,'yyyymm')];      
      
myTitle={'��Ʒ����(1)','��Ʒ��ֵ(2)','��Ʒ�ʲ���ֵ(3)','��Ʒ��λ��ֵ(4)','��Ʒ�ݶ�(5)', ...
'��Ʊ�ʲ���ģ(6)','��֤�����ģ(7)','��֤����Beta(8)','��֤�����ģ(9)', ...
'��������Beta(10)','��С���ģ(11)','��С��Beta(12)','��ҵ���ģ(13)', ...
'��ҵ��Beta(14)','��ļ�����ʲ���ģ(15)','ծȯ�ʲ���ģ(16)','����ծ�ʲ���ģ(17)', ...
'����ծ����(18)','����ծ�ʲ���ģAAA(19)','����ծ�ʲ���ģAA+(20)', ...
'����ծ�ʲ���ģAA(21)','����ծ����AAA(22)','����ծ����AA+(23)','����ծ����AA(24)', ...
'�ֽ��ģ(25)','�ɱ����ʲ���ģ���(26)','�ɱ����ʲ���ģ�ж�(27)','�ɱ����ʲ���ģ�ض�(28)'};

%% ǰ��׼��
files=dir(location);
numPro=0;
namePro=cell(0,1);  % ��Ʒ����
dataPro=cell(0,1);  % ��Ʒ��ֵ������
debtPro=cell(0,1); % ��Ʒծȯ�ֲִ��롢��ơ��ֲֹ�ģ
numWrong=0;
listWrong=cell(numWrong,2);
% ��Ҫ�ȶ�ծȯ��֤ȯ���������ȡ�ͱ��
he=actxserver('Excel.Application');
hw=he.Workbooks.Add;
he.Visible=1;
hs=hw.Worksheets;
sheetItem=hs.Item(1);
for n=1:length(files)
    filename=files(n).name;
    if strfind(filename,'֤ȯͶ�ʻ����ֵ��')        
        numPro=numPro+1;
        disp(['��',num2str(numPro),'ֻ��Ʒ׼����']);
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
        % ����excel��������ȡ֤ȯ���룬�������������
        % 1����׺������������һ�£�2��ֱ�Ӷ�������������������пո�
        indIB=find(cellfun(@(x) str2double(x(1:6))>110340 ,dataPro{numPro}(indDebt,1) ));
        if ~isempty(indIB)
            nameIB=strrep(dataPro{numPro}(indDebt(indIB),2),' ','');
            mvIB=dataPro{numPro}(indDebt(indIB),8);
            codeIB=cell(length(indIB),1);
            for m=1:length(indIB)
                sheetItem.Range('A1').Value=['=to_windcode("',nameIB{m},'")'];
                codeIB{m}=sheetItem.Range('A1').Value;
                if isnumeric(codeIB{m}) % || ~strcmp(codeIB{m}(end-1:end),'IB')   % �����������Ϊ�˾�ȷƥ�����м䣬���������ȡ�۸���ʵ����Ҫ
                    numWrong=numWrong+1;
                    listWrong(numWrong,:)={namePro{numPro},nameIB{m}};
                end
            end
            debtPro{numPro}=cat(1,debtPro{numPro},[codeIB,nameIB,mvIB]);
        end
        xlswrite([location,'\myHold.xlsx'],debtPro{numPro},namePro{numPro});
    end
end
disp(['��',num2str(numWrong),'��ծȯ���������⣡']);

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
        if strfind(filename,'֤ȯͶ�ʻ����ֵ��')            
            rlt=cell(1,length(myTitle));
            ind=strfind(filename,'_');
            rlt{1}=filename(ind(1)+1:ind(2)-1); % ��Ʒ����
            [~,locb]=ismember(rlt{1},namePro);
            myData=dataPro{locb};
            myDebt=debtPro{locb};
            disp([rlt{1},'��ȡ���㣺']);
            
            ind2=find(cellfun(@(x) ischar(x) && ~isempty(strfind(x,'�����ʲ���ֵ')),myData(:,1)));
            rlt{2}=myData{ind2,8};   % ��Ʒ��ֵ
            ind3=find(cellfun(@(x) ischar(x) && ~isempty(strfind(x,'�ʲ���ϼ�')),myData(:,1)));
            rlt{3}=myData{ind3,8};    % �ʲ���ֵ
            ind4=find(cellfun(@(x) ischar(x) && ~isempty(strfind(x,'����λ��ֵ')),myData(:,1)));
            rlt{4}=myData{ind4,2};    % ����λ��ֵ
            ind5=find(cellfun(@(x) ischar(x) && ~isempty(strfind(x,'ʵ���ʱ����')),myData(:,1)));
            rlt{5}=myData{ind5,8};    % �����Ʒ�ݶ�
            ind6=find(cellfun(@(x) ischar(x) && strcmp(x,'1102'),myData(:,1)));
            if ~isempty(ind6)
                rlt{6}=myData{ind6,8};    % ��Ʊ�ʲ���ģ
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
                rlt{7}=sum(mv);     % ��֤�����ģ
                rlt{8}=bet'*mv/sum(mv); % ��֤����Beta
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
                rlt{9}=sum(mv);     % ��֤�����ģ
                rlt{10}=bet'*mv/sum(mv); % ��֤����Beta
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
                rlt{11}=sum(mv);     % ��֤��С���ģ
                rlt{12}=bet'*mv/sum(mv); % ��֤��С��Beta
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
                rlt{13}=sum(mv);     % ��ҵ���ģ
                rlt{14}=bet'*mv/sum(mv); % ��ҵ��Beta
            end
            ind15=find(cellfun(@(x) ischar(x) && length(x)>=14 && strcmp(x(1:4),'1105') && ~isnan(str2double(x(9:end))), myData(:,1) ));
            if ~isempty(ind15)
                rlt{15}=sum(cell2mat(myData(ind15,8)));       % ��ļ�����ʲ���ģ      
            end
            
            ind16=find(cellfun(@(x) ischar(x) && strcmp(x,'1103'),myData(:,1)));
            if ~isempty(ind16)
                rlt{16}=myData{ind16,8};    % ծȯ�ʲ���ģ
                
%                 baseRate=w.wss(myDebt(:,1),'baserate');      %% ��Ϣծ��׼����
                baseRate=w.wsd(myDebt(:,1),'baserate','ED-0D',endT);      %% ��Ϣծ��׼����
                isChange=logical(cellfun(@ischar,baseRate));
%                 dur=w.wss(myDebt(:,1),'modifiedduration',['tradeDate=',endT]);
                dur=w.wsd(myDebt(:,1),'modifiedduration','ED-0D',endT);
                dur(isChange)=0.5;  %% �Ը�Ϣծ����ȡֵΪ0.5
%                 rateDebt=w.wss(myDebt(:,1),'latestissurercreditrating2,rate_ratebond',['tradeDate=',endT],'ratingAgency=101','type=1');   %% ���塢ծ������
                tmp1=w.wsd(myDebt(:,1),'latestissurercreditrating2','ED-0D',endT,'ratingAgency=101','type=1');
                tmp2=w.wsd(myDebt(:,1),'rate_ratebond','ED-0D',endT,'ratingAgency=101','type=1');
                rateDebt=cat(2,tmp1,tmp2);
                
%                 issuer=w.wss(myDebt(:,1),'issuerupdated'); %% ��������
                issuer=w.wsd(myDebt(:,1),'issuerupdated','ED-0D',endT); %% ��������
                indRat=find(cellfun(@(x) ismember(x,{'���ҿ�������';'�й�ũҵ��չ����';'�й�����������';'�л����񹲺͹�������'}) || strcmp(x(end-3:end),'��������') ,issuer)); %% ����ծ�±�
                if ~isempty(indRat)
                    mv=cell2mat(myDebt(indRat,3));
                    rlt{17}=sum(mv);    % ����ծ�ʲ���ģ
                    rlt{18}=dur(indRat)'*mv/sum(mv);    % ����ծ����
                end
                indCre=setdiff(1:size(myDebt,1),indRat)';   %% ����ծ�±�
                if ~isempty(indCre)
                    rate=rateDebt(indCre,:);
                    ind=cellfun(@isnumeric,rate(:,2));
                    rate(ind,2)=rate(ind,1);
                    rate(:,1)=[];
                    
                    ind3A=find(strcmp(rate,'AAA'));
                    if ~isempty(ind3A)
                        mv3A=cell2mat(myDebt(indCre(ind3A),3));
                        rlt{19}=sum(mv3A);      % AAA����ծ��ģ
                        rlt{22}=dur(indCre(ind3A))'*mv3A/sum(mv3A); % AAA����ծ����
                    end
                    ind2AP=find(strcmp(rate,'AA+'));
                    if ~isempty(ind2AP)
                        mv2AP=cell2mat(myDebt(indCre(ind2AP),3));
                        rlt{20}=sum(mv2AP); % AA+����ծ��ģ
                        rlt{23}=dur(indCre(ind2AP))'*mv2AP/sum(mv2AP);  % AA+����ծ����
                    end
                    ind2A=setdiff(1:length(indCre),union(ind3A,ind2AP))';
                    if ~isempty(ind2A)
                        mv2A=cell2mat(myDebt(indCre(ind2A),3));
                        rlt{21}=sum(mv2A);  % AA����������ծ��ģ
                        rlt{24}=dur(indCre(ind2A))'*mv2A/sum(mv2A); % AA����������ծ����
                    end
                end
            end
            ind25=cellfun(@(x) ischar(x) && ismember(x,{'1002','1021','1031'}),myData(:,1));
            rlt{25}=sum(cell2mat(myData(ind25,8)));   % �ֽ��ģ
            
            [~,locb]=ismember(rlt{1},openNext(:,1));
            if locb         %% �������180����δ�������п��ŵļ���ɱ����ʲ�
                dateNext=openNext{locb,2};  %% ��һ������
                numD=double(w.tdayscount(endT,dateNext)-1);
                % �ɱ����ʲ���Ϊ4��
                mv1=rlt{25};
                mv2=0;
                if ~isempty(ind16)
%                     dateMat=w.wss(myDebt(:,1),'maturitydate');
                    dateMat=w.wsd(myDebt(:,1),'maturitydate','ED-0D',endT);
                    indEarly=find(datenum(dateMat)<datenum(dateNext));
                    if ~isempty(indEarly)
                        mv1=mv1+sum(cell2mat(myDebt(indEarly,3)));  % 1������ǰ���ڽ����ʲ��������ֽ�
                    end
                    
                    indLate=setdiff(1:size(myDebt,1),indEarly)';
                    if ~isempty(indLate)
%                         typeHold=w.wss(myDebt(indLate,1),'windl2type');
                        typeHold=w.wsd(myDebt(indLate,1),'windl2type','ED-0D',endT);
                        ind1=find(ismember(typeHold,{'��ծ','��������ծ','����Ʊ��'}));
                        rate=rateDebt(indLate,1);
                        ind2=find(cellfun(@(x) ~isempty(strfind(x,'��������')) ,typeHold) & strcmp(rate,'AAA') );
                        mv2=mv2+sum(cell2mat(myDebt(indLate(union(ind1,ind2)),3))); % 2������ǰδ���ڵ����ڿɱ��ֵ�ծȯ
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
                    mv3=mv3+sum(min(repmat(mvHold,1,3),(hsl.*mvFree)*(1+degreeD)*0.05*numD ));% 3����Ʊ�ɱ��ֲ��� 
                end
                
                mv4=0;
                ind=find(cellfun(@(x) ischar(x) && length(x)>8 && strcmp(x(1:8),'11090301') ,myData(:,1)));
                if ~isempty(ind)
                    mv4=mv4+sum(cell2mat(myData(ind,8)));   % 4����ļ�ʽ��ֽ�����Ʋ�Ʒ
                end
                if ~isempty(rlt{15})
                    mv4=mv4+rlt{15};
                end
                rlt{26}=mv1+mv2+mv3(1)+mv4; % ��ȿɱ����ʲ���ģ
                rlt{27}=mv1+mv2+mv3(2)+mv4; % �жȿɱ����ʲ���ģ
                rlt{28}=mv1+mv2+mv3(3)+mv4; % �ضȿɱ����ʲ���ģ
            end
            result=cat(1,result,rlt);
        end
    end
    
    result=cat(1,myTitle,result);
    xlswrite([location,'\myHold.xlsx'],result,'����');
end