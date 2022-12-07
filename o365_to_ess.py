import csv


def re_organize(raw: str, li: list):
    sender = ''
    for char in raw:
        if char == ' ':
            li.append(sender)
            sender = ''
            continue
        else:
            sender += char


def set_emails(in_list: list, action: str):
    global email_out_list
    for i in in_list:
        temp_list = []
        temp_list.append(i)
        temp_list.append(action)
        temp_list.append('From O365')
        email_out_list.append(temp_list)


def set_ips(in_list: list, action: str):
    global ip_out_list
    for i in in_list:
        temp_list = []
        temp_list.append(i)
        temp_list.append('255.255.255.255')
        temp_list.append(action)
        temp_list.append('From O365')


folder = 'C:\\temp\\'
out_file = f'{folder}email_ess.csv'
out_file2 = f'{folder}ip_ess.csv'

allowed_senders_raw = 'danielle@daiglecreative.com joe.sciara@22genes.com notifications@alarm.com report.builder@flashtalking.com ronreach@gmail.com suzanne@drumwrightmedia.com us-invoice-internal-mbox@oath.com eric@renaissancemarketingva.com att.toddconner9@gmail.com pjsudhakar13@gmail.com yaamanx661@gmail.com'
allowed_domains_raw = ' 1797creative.com 360i.com 3sixtyreach.com 8451.com abelsontaylor.com abovead.com AcrobatAnt.com adfarm.com adtelligent.com affirmagency.com agorafinancial.com americansignature.com apptv.com bbigcommunications.com bearagencygroup.com bionic-ads.com blakejarrett.ca bluepop.digital bohanideas.com bozell.com brandstarbeacon.com brasco.marketing burrell.com c1-partners.com callawaybank.com camelotsmm.com cashtime.com certifications.thetradedesk.com chappellroberts.com chernoffnewman.com commitagency.com cq-media.com creativespot.com culturespanmarketing.com cuneoadvertising.com ddb.com degdigital.com deutschinc.com digitalriver.com divisiond.com echostor.com email.rd.com envoyinc.com epsilon.com ervinandsmith.com estespr.com exlmedia.com fcb.com flashtalking.com flightpathagency.com foursquare.com gelia-media.com goodwaygroup.com gordleygroup.com gudmarketing.com imageadagency.com integrateagency.com intoxalock.com jamesandmatthew.com louisvilleky.gov luum.com marlinnetwork.com mecglobal.com mediapartners-inc.com MediaPost.com methodgroupe.com mferentals.com miamidda.com mni.com modop.com omgblog.com orange142.com paradigmdigital.com performancefirstdigital.com pinnacle-advertising.com placed.com powercreative.com proof-advertising.com ptarmiganmedia.com publicis.com purplestrategies.com rachaelpiperconsulting.com refinitiv.com responsemedia.com rialtosquare.com root3marketing.com salesforce.com sellsagency.com sherrymatthews.com shiftnow.com sifox.net spartanbrandingco.com splunk.com springserve.com staradvertiser.com starcomww.com steelwagon.com tambourine.com targetmktng.com theadcomgroup.com thinkbluemarketing.com thinpigmedia.com thisisaoa.com thomasarts.com tinyhorse.com tipalti.com touchpoint.net tractionfactory.com trefoilgroup.com tuliptree-studios.com uwginc.com vaughnharlow.com velaagency.com vertical3media.com vgroupholdings.com vistarmedia.com vmlyr.com voterx.com wearethirdear.com wmglobal.com wtads.com zedo.com zenogroup.com zoommedia.com bd.com mypointcu.com fast-trackmarketing.com pepsico.com'
blocked_senders_raw = 'charleseddie764@gmail.com kim@thedatasouk.com karabell@karabellindustries.com Jacob.Vaughn@bd.com hersuimichiteio7@gmx.net AdmiT1013@email.com afdzxcs@gmail.com 21c-prod@baswarepm.com sonjal@terrebonneports.com mrodrigue@tmlsupply.com charlesarmstron63@gmail.com jnolan@bairesdev.net foxbasealpha80@gmail.com resankaranco1450@gmail.com contact@kycdepo.com arun.singh@vehere.com'
blocked_domains_raw = 'content.ad whatjobs.com sapo.pt wadax.ne.jp gmx.net'
blocked_ips_raw = '203.183.42.114 212.227.126.130 211.1.227.4 45.156.22.187 81.17.30.239 210.189.85.2 219.118.72.112 210.189.85.18 153.127.234.5 140.227.179.114 153.127.234.174 153.127.234.79 153.127.234.232 133.242.55.7 153.127.234.3 153.127.234.177 211.1.224.234 217.67.28.16 93.184.77.214 40.107.117.107'

allowed_senders_list = []
allowed_domains_list = []
blocked_senders_list = []
blocked_domains_list = []
blocked_ips_list = []

email_headers = ['Email Address', 'Policy (block, exempt, quarantine)', 'Comment (optional)']
ip_headers = ['IP Address', 'Netmask', 'Policy (block, exempt, quarantine)', 'Comment (optional)']


re_organize(allowed_senders_raw, allowed_senders_list)
re_organize(allowed_domains_raw, allowed_domains_list)
re_organize(blocked_senders_raw, blocked_senders_list)
re_organize(blocked_domains_raw, blocked_domains_list)
re_organize(blocked_ips_raw, blocked_ips_list)

email_out_list = []
ip_out_list = []
set_emails(allowed_senders_list, 'exempt')
set_emails(allowed_domains_list, 'exempt')
set_emails(blocked_senders_list, 'blocked')
set_emails(blocked_domains_list, 'blocked')
set_ips(blocked_ips_list, 'blocked')

with open(out_file, 'w', encoding='utf-8', newline='') as output_file:
    writer = csv.writer(output_file, quoting=csv.QUOTE_NONNUMERIC)
    writer.writerow(email_headers)
    writer.writerows(email_out_list)

with open(out_file2, 'w', encoding='utf-8', newline='') as output_file:
    writer = csv.writer(output_file, quoting=csv.QUOTE_NONNUMERIC)
    writer.writerow(ip_headers)
    writer.writerows(ip_out_list)

