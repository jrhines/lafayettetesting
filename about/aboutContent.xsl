<?xml version="1.0" encoding="ISO-8859-1"?>

<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
xmlns="urn:schemas-microsoft-com:xml-data" 
xmlns:dt="urn:schemas-microsoft-com:datatypes"
xmlns:msxsl="urn:schemas-microsoft-com:xslt"              
xmlns:user="Deals"
extension-element-prefixes="msxsl user">

<xsl:output method="html" />

<xsl:param name="selectedContent" />
<xsl:param name="XMLFile" />
<xsl:param name="image" />

<xsl:template match="/">
	<xsl:apply-templates select="contents/content"/>
</xsl:template>

<xsl:template match="contents/content">
	<h1 id="headline" class="margin_10_0_0_0"><xsl:value-of select="headline" disable-output-escaping="yes"/></h1>
	<xsl:choose>
		<xsl:when test="($selectedContent = 'overview') or ($selectedContent = 'lab_diversification')">
			<xsl:for-each select="paragraph">
				<p><xsl:value-of select="." disable-output-escaping="yes"/></p>
			</xsl:for-each>
		</xsl:when>
		<xsl:when test="$selectedContent = 'image_gallery'">
			<xsl:choose>
				<xsl:when test="$image = ''"> 
					<xsl:apply-templates select="gallery/image"/>
				</xsl:when>
				<xsl:otherwise>
					<img name="xsl" border="1" width="400" class="margin_0_10_10_0">
						<xsl:attribute name="src">../images/about/images/<xsl:value-of select="$image"/>.jpg</xsl:attribute>
						<xsl:attribute name="alt"><xsl:value-of select="gallery/image[@alt = $image]" disable-output-escaping="yes" /></xsl:attribute>
					</img>
					<p><xsl:value-of select="gallery/image[@alt = $image]/info"/></p>
					<p class="fineprint"><a><xsl:attribute name="href">index.asp?selectedContent=<xsl:value-of select="$selectedContent"/>&amp;XMLFile=<xsl:value-of select="$XMLFile"/></xsl:attribute>Return to Thumbnail View</a></p>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:when>
	</xsl:choose>
</xsl:template>

<xsl:template match="gallery/image">
	<a><xsl:attribute name="href">index.asp?selectedContent=<xsl:value-of select="$selectedContent"/>&amp;XMLFile=<xsl:value-of select="$XMLFile"/>&amp;image=<xsl:value-of select="@alt"/></xsl:attribute>
	<img name="xsl" border="0" width="90" class="margin_0_2_2_0">
		<xsl:attribute name="src">../images/about/thumbnails/<xsl:value-of select="@alt"/><xsl:value-of select="file"/></xsl:attribute>
		<xsl:attribute name="alt"><xsl:value-of select="@alt" disable-output-escaping="yes" /></xsl:attribute>
		<xsl:attribute name="title"><xsl:value-of select="Info" disable-output-escaping="yes" /></xsl:attribute>
	</img></a>						
</xsl:template>

</xsl:stylesheet>