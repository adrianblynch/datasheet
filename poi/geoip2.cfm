
<cfset userID = "81970">
<cfset key = "GJOsS0E4pYBh">
<cfset ip = "85.133.39.128">

<cfhttp url="https://geoip.maxmind.com/geoip/v2.0/country/#ip#">
	<cfhttpparam type="header" name="Authorization" value="Basic #userID#:#key#">
</cfhttp>

<cfdump var="#CFHTTP#" label="CFHTTP">

<!--- // This creates a WebServiceClient object that can be reused across requests.
// Replace "42" with your user ID and "license_key" with your license key.
WebServiceClient client = new WebServiceClient.Builder(42, "license_key").build();

// Replace "omni" with the method corresponding to the web service that
// you are using, e.g., "country", "cityIspOrg", "city".
OmniResponse response = client.omni(InetAddress.getByName("128.101.101.101"));

System.out.println(response.getCountry().getIsoCode()); // 'US'
System.out.println(response.getCountry().getName()); // 'United States'
System.out.println(response.getCountry().getNames().get("zh-CN")); // '美国'

System.out.println(response.getMostSpecificSubdivision().getName()); // 'Minnesota'
System.out.println(response.getMostSpecificSubdivision().getIsoCode()); // 'MN'

System.out.println(response.getCity().getName()); // 'Minneapolis'

System.out.println(response.getPostal().getCode()); // '55455'

System.out.println(response.getLocation().getLatitude()); // 44.9733
System.out.println(response.getLocation().getLongitude()); // -93.2323 --->