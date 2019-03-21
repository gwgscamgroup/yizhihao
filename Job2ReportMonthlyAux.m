% ����Job2��������Ʒ�ֲ��±���д
%% ����
location='D:\Job2\�ֲ��±�201902';
%%
files=dir(location);
if exist([location,'\myHold.xlsx'],'file')
    % ��ȷ����ɵ�һ������ȷ�ϴ��������ظ�
    numWrong=0;
    for n=1:length(files)
        filename=files(n).name;
        if strfind(filename,'֤ȯͶ�ʻ����ֵ��_');
            ind=strfind(filename,'_');
            namePro=filename(ind(1)+1:ind(2)-1);
%             if strfind(namePro,'����1��')
%                 continue;
%             end
            namePro=namePro(7:10);
            [~,~,myHold]=xlsread([location,'\myHold.xlsx'],namePro);
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
        for n=1:length(files)
            filename=files(n).name;
            if strfind(filename,'֤ȯͶ�ʻ����ֵ��_');
                ind=strfind(filename,'_');
                namePro=filename(ind(1)+1:ind(2)-1);
%                 if strfind(namePro,'����1��')
%                     continue;
%                 end
                namePro=namePro(7:10);
                [~,~,myHold]=xlsread([location,'\myHold.xlsx'],namePro);
                numHold=size(myHold,1);
                holdInfor=cell(numHold,6);
                holdInfor(:,1:4)=myHold(:,1:4);
%                 holdInfor(:,5)=w.wss(myHold(:,1),'latestissurercreditrating');  % ��������
%                 holdInfor(:,6)=w.wss(myHold(:,1),'amount');     % ծ������
%                 holdInfor(:,7)=w.wss(myHold(:,1),'industry_sw','industryType=1');
                holdInfor(:,5)=w.wsd(myHold(:,1),'latestissurercreditrating');  % ��������
                holdInfor(:,6)=w.wsd(myHold(:,1),'amount');     % ծ������
                holdInfor(:,7)=w.wsd(myHold(:,1),'industry_sw','','','industryType=1');
                xlswrite([location,'\myHold.xlsx'],holdInfor,namePro);
                % ����������ծ������ʱ�����������������棩
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
                xlswrite([location,'\myHold.xlsx'],grade,[namePro,'����']);
                % ��ҵ����
                holdTmp=holdInfor(~cellfun(@isnumeric,holdInfor(:,7)),:);
                valSum=sum(cell2mat(holdTmp(:,4)));
                uni=unique(holdTmp(:,7));
                prob=nan(length(uni),1);
                for i=1:length(uni)
                    ind=find(strcmp(holdTmp(:,7),uni{i}));
                    prob(i)=sum(cell2mat(holdTmp(ind,4)))/valSum;
                end
                idy=cat(2,uni,num2cell(prob));
                xlswrite([location,'\myHold.xlsx'],idy,[namePro,'��ҵ']);
            end
        end
        
    end
else
    % ��һ�����Ƚ�������Ʒ��ծȯ�ֲּ���ȡ�����ֶ�����wind����
    % �ֶ�����ⲽ���ٴ����д��ļ�
    for n=1:length(files)
        filename=files(n).name;
        if strfind(filename,'֤ȯͶ�ʻ����ֵ��_')
            ind=strfind(filename,'_');
            namePro=filename(ind(1)+1:ind(2)-1);
%             if strfind(namePro,'����1��')
%                 continue;
%             end
            namePro=namePro(7:10);
            [~,~,data]=xlsread([location,'\',filename]);
            indDebt=cellfun(@(x) ischar(x) && length(x)>4 && (strcmp(x(1:4),'1103') || strcmp(x(1:4),'1104')),data(:,1) );
            indValid=cellfun(@(x) isnumeric(x) && ~isnan(x),data(:,4));
            indChs=indDebt & indValid;
            myHold=data(indChs,[2,3,8,5,1]);% �ֱ���ȡ��ơ���������ֵ���ɱ�����Ŀ����
            if strfind(namePro,'����1��')
                myHold=cat(1,myHold,{'16����03',500000,50000000,50000000,0});
            end            
            xlswrite([location,'\myHold.xlsx'],myHold,namePro);
        end
    end
    disp('�Ѽ���ȡ��ծȯ��ƣ��ֶ�����wind��������ٴ����б��ļ�');
end