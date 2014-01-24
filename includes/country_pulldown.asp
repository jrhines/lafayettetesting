<select name="country">
	<option value="US" <%if Request("Country") = "US" then Response.Write( "selected" ) end if%>>USA</option>
	<option value="CA" <%if Request("Country") = "CA" then Response.Write( "selected" ) end if%>>Canada</option>
	<option value="MX" <%if Request("Country") = "MX" then Response.Write( "selected" ) end if%>>Mexico</option>
	<option value="UK" <%if Request("Country") = "UK" then Response.Write( "selected" ) end if%>>UK</option>
	<option value="OT" <%if Request("Country") = "OT" then Response.Write( "selected" ) end if%>>Other</option>
</select>