from DCNM_Authentication import headers_token,dcnm_ip,fabric_name

url_vrf_attach = f"https://{dcnm_ip}/rest/top-down/fabrics/{fabric_name}/vrfs/attachments"