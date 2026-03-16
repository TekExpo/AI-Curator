# Connect (delegated)
Connect-MgGraph -Scopes "Group.ReadWrite.All", "User.Read"

# Get token (MSAL way or from Connect-MgGraph)
$token = (Get-MgContext).AuthToken  # or use Get-MsalToken for more control

$token = '1.AXwAsGn19cU750ybRgKPo7sAO-wu2BRLIC9Mt-gpanDatn4aAA18AA.BQABBAIAAAADAOz_BQD0_0V2b1N0c0FydGlmYWN0cwIAAAAAALMQlUzf6ajEluk8l7QxKMem29AJJ1ifXuzqhcMWbh6MkSUfpECkKXnAR_QZVKbWX6T86uA9iYVetO4EP7C4RexzBnQrWi37SmQy2BVYwKDOHz92HwGz8yObzgLeMQlUStf0FVVDxj9DX0RXCJa7Qhw8zP98_mhvJnlcKvaY39gI2JLm5-scZPlXvdxSZHNvPHcv1KrLdbVPoyy9FUrCKb8LK2wYONRuHHaZx5lQxwgZyWKqQM3W5_RpUocFmeiB4rOviGC57CwNyLH79dGLkz_VSbwIeoRjIqGvWUrjUyTETy8zf5y4Mi5WkgR-ho_FkloI9f1A-wDRFnfvFvEtBTRD6Ou8vcWC3PaAxfG5PwK-DBlyM1DYu_-8BL2FdopxQg51IF4pQFY9DOXEjJwxwZmEMW-ykiUMZUth6OrM3f7yriW6s47GYWUvTvth1UTnaR4WhJzqeBCskR4sOYrqN1ejGpbptIT9GWjz-LsdVzquvI42ZdajhE11F1fCmUeHmq_tDzLyKRZ-Gm37CMrRC23GbXnAAxMvyY1kH_SojCtHpp_XmC8UrCYU6W4kfrJL8s0EcXlfN4UDof3CbOFx2Agax5Jh5wW2ta_DxtucNeOFAhseeRLtbBvZTVYgpxp3jHSCXJX7ol4CpqnINPpcSJrcQyK6aY53cm4niO7GWhXP1Es9a4FEorZym0tubvhKpIfvdrMt0pzyHtfKlVzMlK6xQBfrWKDZ4qCdOly_bAu_ccp-SlVNQ7gV4lGAEb3tSPJdbTaU0rf4Mm3a5ur4oSJnN6o6PIVXFLc3ox_GpVRl_3YzHyMcOL-UxED4u3xpD0gpwjSPboDmUzQ2TzJwncUNEDwEiEeNbj86EwUmiu9V6GRpYKb0e5YGDh2Dc4ON9gw1FQ8y0itPX4A3e59QeUfkpC7MyUzrU0ym0_JWJIXxXQUEWCArQZTy7vdpmK083ofWuJsB7JkuBNqKtPH9qpPcEm-KzLdqJzsm3i2QbEYzBWs5CfQEcryZQ66Z5iG8D5A5zQWx50AWBPH3MoVLj5ihRQHMhAduYFdoKuPztcCSYEGKVBgS_7NT25MZGab4O7CcmrUtlOAOFL9qFthKv27uaGFyFxBdlq9_UQxFItkiZbz19xNM12W828ujvEk6_Qf0e70GGt9CQIlRqfo6J4CZ5oDoFVwA4tBSbO4tDPKGMjAixUsQmQ98SO1OvYl637PerseVaKIj9-oUXL0alhGHSRXK6NjDRSK6-2VpclHP-_Qp-pH2pxcGs32eBdGXDMQW2M_FVebYwAvRx0pdfQWDpg92deSpw83CYJRfmEwCnQ5-TJKV-0OQ0_Db0WI8Lev2Un8TwU9c0RmAC8vJfx7L0d3BXqmk2WXnYwBIk08Cv8SvTmRljZazZz1mMLvVWMWJG7vWH7eiWgsRzg'
$headers = @{
    Authorization = "Bearer $token"
    "Content-Type" = "application/json"
}

$body = @{
    body     = "Test post via legacy API from PowerShell"
    group_id = "1a9ad179-d733-4a98-821a-17bd34cbea8e"
} | ConvertTo-Json

Invoke-RestMethod -Uri "https://www.yammer.com/api/v1/messages.json" `
                  -Method Post `
                  -Headers $headers `
                  -Body $body