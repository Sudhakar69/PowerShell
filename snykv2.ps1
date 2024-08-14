
# Define your Snyk API token
$token = "Snyk API token"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$headers = @{
    Authorization="Bearer $token"
}
# Define the URL for the API request
$url = "https://app.snyk.io/org/$($organization)/projects"
# $url = "https://app.snyk.io/oauth2/authorize"
# $url = "https://app.snyk.io/account/"
# $url = "https://app.snyk.io/org/$($project-org)"
# $url = "https://app.snyk.io/org/$($project-org)/reporting?v=1&context[page]=issues-detail&issue_status=%255B%2522Open%2522%255D&issue_by=Severity&table_issues_detail_cols=SEVERITY%257CSCORE%257CCVE%257CCWE%257CPROJECT%257CEXPLOIT%2520MATURITY%257CAUTO%2520FIXABLE%257CINTRODUCED%257CSNYK%2520PRODUCT"

$responseData = (Invoke-RestMethod -Uri $Url -Method Get -Headers $headers -UseBasicParsing -ContentType "application/json").Content | ConvertFrom-Json | ConvertTo-Json
$responseData