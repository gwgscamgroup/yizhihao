% ͳ�ƹ�Ʊ�ֲ����Ӧ�ĸ��ʱ仯
%% ����
location='D:\Job2\��������\��Ʊ�ֲ�.xlsx';
begT=20190301;
endT=20190320;
%% ���ݶ�ȡ
[~,~,dataTmp]=xlsread(location);
codeTs=CodeW2T(dataTmp(:,1));

ts=actxserver('TSExpert.CoExec');
dataLgt=ts.RemoteCallFunc('pdGangGuChiCang',{codeTs});
myData=dataLgt(2:end,:);
dList=str2num(datestr(ts.RemoteCallFunc('pdTime',{begT,endT,'����'}),'yyyymmdd')); %#ok<ST2NM>
shareStk=ts.RemoteCallFunc('pdTotalShares',{codeTs,dList(1),dList(end)});
shareStk=cell2mat(shareStk(2:end,:));
%% ͳ������
numS=length(codeTs);
numD=length(dList);
sumHold=nan(numD,numS);
for s=1:numS
    if ~isempty(myData{s,3}) && double(myData{s,3}{end,1})>=endT
        tmp=myData{s,3}(2:end,:);
        % ��Ҫע���½��ͨ����������۹��г���Ӧ��
        % ���Խ���������A�ɽ��������в��
        [~,ind]=ismember(double(cell2mat(tmp(:,1))),dList);
        indValid=find(ind>0);
        sumHold(ind(indValid),s)=cell2mat(tmp(indValid,2));
        
        indB=find(~isnan(sumHold(:,s)),1);
        indN=intersect(find(isnan(sumHold(:,s))),indB:numD);
        if ~isempty(indN)
            for n=1:length(indN)
                sumHold(indN(n),s)=sumHold(indN(n)-1,s);
            end
        end
        
    end
end
rate=sumHold./shareStk;
tmp=diff(rate(end-1:end,:),[],1);
[sortR,indSort]=sort(tmp);
rlt=[dataTmp(indSort,:),num2cell(sortR')];