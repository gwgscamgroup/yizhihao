% ����Job2�е��ʹ��ܱ��볤���ܱ���д
%% ����
% ����ʱ���޸�
begT=20190308;
endT=20190315;
useWsd=1;   % �Ƿ�Ҫ��wsd����wssȡ����
% һ�������޸�
location='D:\Job2\��ά�ܱ�';
locHolder='D:\Job2\��������˷ݶ�';
inforPro={'���ǹ���֤ȯ����1�Ŷ����ʲ�����ƻ�','G-B-1','SG8892'; ...
          '���ǹ���֤ȯ����2�Ŷ����ʲ�����ƻ�','G-C-2','SG9617'; ...
          '���ǹ���֤ȯ����1�ż����ʲ�����ƻ�','G-C-1','SQ4475'; ...
          '���ǹ���֤ȯ��ӯ1�Ŷ����ʲ�����ƻ�','G-D-1','SQ0803'; ...
          '���ǹ���֤ȯ����6�Ŷ����ʲ�����ƻ�','G-D-6','SET595'};
inforType={'ͬҵ�浥',0;'��ծ',1;'�ط�����ծ',2;'����Ʊ��',3;'�����Խ���ծ',4;'��������ծ',4;'��ҵ����ծ',5; ...
'��ҵ���дμ�ծ',5;'���չ�˾ծ',5;'֤ȯ��˾ծ',5;'֤ȯ��˾��������ȯ',5;'�������ڻ���ծ',5; ...
'һ����ҵծ',6;'������ҵծ',6;'NPB',6;'һ�㹫˾ծ',7;'����Ʊ��',8;'һ������Ʊ��',8;'һ���������ȯ',9; ...
'����������ծȯ',9;'PPN',10;'���ʻ���ծ',11;'����֧�ֻ���ծ',12;'֤�������ABS',13; ...
'���������ABS',13;'������Э��ABN',13;'��Ŀ����Ʊ��PRN',13;'��תծ',14;'�ɽ���ծ',15; ...
'�ɷ���תծ��ծ',16;'�ǹ������й�˾ծ',17;'˽ļծ',17;'����',18};
%% 
if useWsd
    fun='w.wsd';
else
    fun='w.wss';    %#ok<UNRCH>
end
locLtday=[location,num2str(begT),'\��3-2��CISP�ʹ��ܱ���',num2str(begT),'��Ͷ�ܣ�.xlsx'];
locToday=[location,num2str(endT)];
files=dir(locToday);
if exist([locToday,'\myHold.xlsx'],'file')
    [~,mySheets,~]=xlsfinfo([locToday,'\myHold.xlsx']);
    % ��ȷ����ɵ�һ������ȷ�ϴ��������ظ�
    numWrong=0;
    for n=1:length(files)
        filename=files(n).name;
        if strfind(filename,'֤ȯͶ�ʻ����ֵ��_');
            ind=strfind(filename,'_');
            namePro=filename(ind(1)+1:ind(2)-1);
            if ~ismember(namePro,mySheets)
                continue;
            end
            [~,indHold]=ismember(namePro,inforPro(:,1));
            [~,~,myHold]=xlsread([locToday,'\myHold.xlsx'],namePro);
            % �ȴ������ظ�wind���루��ֻ����ʱ�Ĵ���ʽ���Ժ��ٸ��ƣ�
            [~,~,ic]=unique(myHold(:,1));
            tmp=tabulate(ic);
            indRep=find(ismember(ic,tmp(tmp(:,2)>1,1)));
            for i=1:length(indRep)
                ind=indRep(i);
                if (strcmp(myHold{ind,1}(end-1:end),'IB') && isnumeric(myHold{ind,6})) || ...
                   (~strcmp(myHold{ind,1}(end-1:end),'IB') && ischar(myHold{ind,6}))
                    disp([namePro,'��',num2str(ind),'ֻծȯ�Ĵ��������⣡']);
                    numWrong=numWrong+1;
                end
            end
            
        end
    end
    disp(['��',num2str(numWrong),'ֻծȯ�Ĵ���������']);
    % 
    if ~numWrong
        w=windmatlab;
        [~,~,data]=xlsread(locLtday,'3_5ծȯҵ�񱨱�');   % ��ȡ��һ�ڵĳֲ�������Ϊɸѡ�䶯��Ķ���
        data([1:4,end],:)=[];
        data(:,1)=[];
        
        result=cell(0,15);
        for n=1:length(files)
            filename=files(n).name;
            if strfind(filename,'֤ȯͶ�ʻ����ֵ��_');
                ind=strfind(filename,'_');
                namePro=filename(ind(1)+1:ind(2)-1);
                if ~ismember(namePro,mySheets)
                    continue;
                end
                [~,indHold]=ismember(namePro,inforPro(:,1));
                [~,~,myHold]=xlsread([locToday,'\myHold.xlsx'],namePro);
                numHold=size(myHold,1);
                rltTmp=cell(numHold,15);
                rltTmp(:,1:3)=repmat(inforPro(indHold,[2,1,3]),numHold,1);  % �ֱ��Ǳ�š����ơ���Ʒ����
                eval(['typeHold=',fun,'(myHold(:,1),''windl2type'');']);
%                 typeHold=w.wss(myHold(:,1),'windl2type');
                [~,indType]=ismember(typeHold,inforType(:,1));
                indType(indType==0)=size(inforType,1);   % ����Ҫ�ر�ע����Ϊ18�ģ���������������ƥ����ֶδչ�����
                rltTmp(:,4)=inforType(indType,2);   % ծȯ���ͱ���
                rltTmp(:,5)=myHold(:,2);    % ծȯ���
                rltTmp(:,6)=myHold(:,1);    % ծȯ����
                eval(['rltTmp(:,7)=',fun,'(myHold(:,1),''carrydate'');']);
                eval(['rltTmp(:,8)=',fun,'(myHold(:,1),''maturitydate'');']);
                eval(['rltTmp(:,9)=',fun,'(myHold(:,1),''amount'');']);
%                 rltTmp(:,7)=w.wss(myHold(:,1),'carrydate'); % ��Ϣ��
%                 rltTmp(:,8)=w.wss(myHold(:,1),'maturitydate');  % ������
%                 rltTmp(:,9)=w.wss(myHold(:,1),'amount');    % ծ������
                rltTmp(cellfun(@isnumeric,rltTmp(:,9)),9)={'��'};
                eval(['rltTmp(:,10)=',fun,'(myHold(:,1),''latestissurercreditrating'');']);
%                 rltTmp(:,10)=w.wss(myHold(:,1),'latestissurercreditrating');    % ��������
                rltTmp(cellfun(@isnumeric,rltTmp(:,10)),10)={'��'};
                rltTmp(:,11:12)=myHold(:,4:5);  % ��ֵ��ɱ�
                eval(['rltTmp(:,13)=num2cell(',fun,'(myHold(:,1),''couponrate2'')/100);']);
%                 rltTmp(:,13)=num2cell(w.wss(myHold(:,1),'couponrate2')/100);  % Ʊ������
                rltTmp(:,14)=repmat({'-'},numHold,1);   % ��ע
                [~,indBR]=ismember(myHold(:,1),{'101778002.IB';'112494.SZ';'136388.SH'});
                rltTmp(indBR==1,14)={'ΥԼ'};
                rltTmp(indBR>=2,14)={'ΥԼ'};
                
                if strfind(namePro,'����1��')
              %% ���������ί���˵��ܱ�
                    holdCR=cell(numHold,6);
                    holdCR(:,1:4)=myHold(:,1:4);
                    holdCR(:,5)=rltTmp(:,10);
                    if useWsd
                        holdCR(:,6)=w.wsd(myHold(:,1),'industry_sw','','','industryType=1');
                    else
                        holdCR(:,6)=w.wss(myHold(:,1),'industry_sw','industryType=1'); %#ok<UNRCH>
                    end                    
                    xlswrite([locToday,'\myHold.xlsx'],holdCR,'����1��ծȯ�ֲ�');
                    % ��������
                    holdTmp=holdCR(~strcmp(holdCR(:,5),'��'),:);
                    valSum=sum(cell2mat(holdTmp(:,4)));
                    uni=unique(holdTmp(:,5));
                    prob=nan(length(uni),1);
                    for i=1:length(uni)
                        ind=find(strcmp(holdTmp(:,5),uni{i}));
                        prob(i)=sum(cell2mat(holdTmp(ind,4)))/valSum;
                    end
                    grade=cat(2,uni,num2cell(prob));
                    xlswrite([locToday,'\myHold.xlsx'],grade,'����1������ռ��');
                    % ��ҵ����
                    holdTmp=holdCR(~cellfun(@isnumeric,holdCR(:,6)),:);
                    valSum=sum(cell2mat(holdTmp(:,4)));
                    uni=unique(holdTmp(:,6));
                    prob=nan(length(uni),1);
                    for i=1:length(uni)
                        ind=find(strcmp(holdTmp(:,6),uni{i}));
                        prob(i)=sum(cell2mat(holdTmp(ind,4)))/valSum;
                    end
                    idy=cat(2,uni,num2cell(prob));
                    xlswrite([locToday,'\myHold.xlsx'],idy,'����1����ҵռ��');
                end
                indCd=find(indType==1);
                if ~isempty(indCd)
                    disp([namePro,'��',rltTmp(indCd,5)]);
                    rltTmp(indCd,:)=[];
                end
                result=cat(1,result,rltTmp);    % �ʹ��ܱ�ծ��ֲ��ﲻ��ͬҵ�浥
            end
        end
        % ɸѡ���б䶯����Ŀ
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
        % ����1�ų�������Ϣ
        fileHolder=dir(locHolder);
        [~,indMax]=max(arrayfun(@(x) x.datenum,fileHolder(3:end)));
        [~,~,tmp]=xlsread([locHolder,'\',fileHolder(indMax+2).name]);
        nameHolder=tmp(2:end-1,2);
        numHolder=cell2mat(tmp(2:end-1,9));
        [numSort,indSort]=sort(numHolder,'descend');
        probHolder=cat(2,nameHolder(indSort),num2cell([numSort,numSort/sum(numSort)]));
        xlswrite([locToday,'\myHold.xlsx'],probHolder,'�����˷ݶ�');
        
        % ������н��
        xlswrite([locToday,'\myHold.xlsx'],result,'result');
    end
else
    % ��һ�����Ƚ�������Ʒ��ծȯ�ֲּ���ȡ�����ֶ�����wind����
    % �ֶ�����ⲽ���ٴ����д��ļ�
    for n=1:length(files)
        filename=files(n).name;
        if strfind(filename,'֤ȯͶ�ʻ����ֵ��_');
            ind=strfind(filename,'_');
            namePro=filename(ind(1)+1:ind(2)-1);
            [~,~,data]=xlsread([locToday,'\',filename]);
            indDebt=cellfun(@(x) ischar(x) && length(x)>4 && (strcmp(x(1:4),'1103') || strcmp(x(1:4),'1104')),data(:,1) );
            indValid=cellfun(@(x) isnumeric(x) && ~isnan(x),data(:,4));
            indChs=indDebt & indValid;
            myHold=data(indChs,[2,3,8,5,1]);% �ֱ���ȡ��ơ���������ֵ���ɱ�����Ŀ����
            
%             indS=find(cellfun(@(x) strcmp(x,'ծȯͶ��'),data(:,2)));
%             indE=find(cellfun(@(x) length(x)>2 && strcmp(x(1:2),'12'),data(:,1)),1);
%             indValid=find(cellfun(@(x) isnumeric(x) && ~isnan(x),data(:,4)));
%             indChs=find(indValid>indS & indValid<indE);
%             myHold=data(indValid(indChs),[2,3,8,5,1]);  % �ֱ���ȡ��ơ���������ֵ���ɱ�����Ŀ����
            if strfind(namePro,'����1��')
%                 myHold=cat(1,myHold,{'16����03',500000,50000000,50000000,0;'16����04',70000,7000000,7000000,0;'17��̩��ԴMTN001',700000,70000000,70000000,0});
                myHold=cat(1,myHold,{'16����03',500000,50000000,50000000,0});
            end
            if ~isempty(myHold)
                xlswrite([locToday,'\myHold.xlsx'],myHold,namePro);
            end
        end
    end
    disp('�Ѽ���ȡ��ծȯ��ƣ��ֶ�����wind��������ٴ����б��ļ�');
end