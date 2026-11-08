#!/usr/bin/env bash
# Probe a SharePoint 2019 / 2016 site with the REST and client object model calls sp2lr relies on.
# Prints one line per check; compare with docs/TESTING-SHAREPOINT-2016-2019.md.
#
#   docker/sharepoint-mock/check.sh 8093      # SharePoint 2019 mock
#   docker/sharepoint-mock/check.sh 8094      # SharePoint 2016 mock
#
# Needs curl (with NTLM support) and python3. Read-only: the only POSTs are /_api/contextinfo
# and a ProcessQuery that reads a term set.
P=${1:?usage: check.sh <port>}; SITE=http://localhost:$P/sites/marketing; API=$SITE/_api
AUTH=(--ntlm -u 'CONTOSO\spadmin:Passw0rd!'); J=(-H 'Accept: application/json;odata=nometadata')
code() { curl -s -o /dev/null -w '%{http_code}' "$@"; }
echo "1 sign-in: none=$(code $API/web) wrong-pw=$(code --ntlm -u 'CONTOSO\spadmin:nope' $API/web) right=$(code "${AUTH[@]}" $API/web)"
echo "2 lists: $(curl -s "${AUTH[@]}" "${J[@]}" "$API/web/lists?\$select=Title,BaseTemplate" | python3 -c "import json,sys;print([(l['Title'],l['BaseTemplate']) for l in json.load(sys.stdin)['value']])")"
SPID=dddddddd-0000-0000-0000-000000000001
echo "3 Site Pages has CanvasContent1 field: $(curl -s "${AUTH[@]}" "${J[@]}" "$API/web/lists(guid'$SPID')/fields?\$select=InternalName" | python3 -c "import json,sys;print('CanvasContent1' in [f['InternalName'] for f in json.load(sys.stdin)['value']])")"
echo "4 \$select=CanvasContent1 on Site Pages: HTTP $(code "${AUTH[@]}" "$API/web/lists(guid'$SPID')/items?\$select=Id,CanvasContent1")"
echo "5 pages: $(curl -s "${AUTH[@]}" "${J[@]}" "$API/web/lists(guid'$SPID')/items?\$select=Id,FileLeafRef,WikiField&\$top=10" | python3 -c "
import json,sys;d=json.load(sys.stdin);print([(i['FileLeafRef'],'wiki' if i.get('WikiField') else 'modern') for i in d['value']], '| nextLink:', 'odata.nextLink' in d)")"
echo "6 publishing page: $(curl -s "${AUTH[@]}" "${J[@]}" "$API/web/lists(guid'dddddddd-0000-0000-0000-000000000002')/items?\$select=FileLeafRef,PublishingPageContent" | python3 -c "import json,sys;i=json.load(sys.stdin)['value'][0];print(i['FileLeafRef'], i['PublishingPageContent'][:40])")"
echo "7 term store REST v2.1: HTTP $(code "${AUTH[@]}" "$API/v2.1/termStore/sets/cccccccc-0000-0000-0000-000000000001")"
DIGEST=$(curl -s "${AUTH[@]}" "${J[@]}" -X POST $API/contextinfo | python3 -c "import json,sys;print(json.load(sys.stdin)['FormDigestValue'])")
Q='<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="test"><Actions><Query Id="11" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="Id" ScalarProperty="true"/><Property Name="Name" ScalarProperty="true"/></Properties></Query></Query><Query Id="12" ObjectPathId="7"><Query SelectAllProperties="false"><Properties/></Query><ChildItemQuery SelectAllProperties="false"><Properties><Property Name="Id" ScalarProperty="true"/><Property Name="Name" ScalarProperty="true"/><Property Name="PathOfTerm" ScalarProperty="true"/></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="1" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}"/><Method Id="3" ParentId="1" Name="GetDefaultSiteCollectionTermStore"/><Method Id="5" ParentId="3" Name="GetTermSet"><Parameters><Parameter Type="Guid">{cccccccc-0000-0000-0000-000000000001}</Parameter></Parameters></Method><Method Id="7" ParentId="5" Name="GetAllTerms"/></ObjectPaths></Request>'
echo "8 CSOM without digest: $(curl -s "${AUTH[@]}" -H 'Content-Type: text/xml' -X POST $SITE/_vti_bin/client.svc/ProcessQuery -d "$Q" | python3 -c "import json,sys;print(json.load(sys.stdin)[0]['ErrorInfo']['ErrorMessage'])")"
echo "9 CSOM with digest: $(curl -s "${AUTH[@]}" -H 'Content-Type: text/xml' -H "X-RequestDigest: $DIGEST" -X POST $SITE/_vti_bin/client.svc/ProcessQuery -d "$Q" | python3 -c "import json,sys;r=json.load(sys.stdin);print([t['PathOfTerm'] for t in r[4]['_Child_Items_']])")"
echo "10 users: $(curl -s "${AUTH[@]}" "${J[@]}" "$API/web/siteusers" | python3 -c "import json,sys;print([u['LoginName'] for u in json.load(sys.stdin)['value']])")"
echo "11 PeopleManager: $(curl -s "${AUTH[@]}" "${J[@]}" "$API/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='i%3A0%23.w%7Ccontoso%5Cjane'" | python3 -c "import json,sys;d=json.load(sys.stdin);print(d['DisplayName'], {p['Key']:p['Value'] for p in d['UserProfileProperties']}['SPS-JobTitle'])")"
echo "12 file download: $(curl -s "${AUTH[@]}" "$API/web/GetFileByServerRelativeUrl('/sites/marketing/Policies/readme.txt')/\$value")"
