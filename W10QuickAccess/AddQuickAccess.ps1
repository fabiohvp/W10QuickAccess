Param
  (
     [string]$p=""
  ) 

Write-Output $p
$o = new-object -com shell.application
$o.Namespace($p).Self.InvokeVerb("pintohome")