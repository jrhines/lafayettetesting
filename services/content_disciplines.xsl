<?xml version="1.0" encoding="ISO-8859-1"?>

<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
xmlns="urn:schemas-microsoft-com:xml-data" 
xmlns:dt="urn:schemas-microsoft-com:datatypes"
xmlns:msxsl="urn:schemas-microsoft-com:xslt"              
xmlns:user="Deals"
extension-element-prefixes="msxsl user">

<xsl:output method="html" />

<xsl:param name="serviceType" />
<xsl:param name="selectedContent" />
<xsl:param name="XMLFile" />
<xsl:param name="image" />

<xsl:template match="/">
	<xsl:apply-templates select="Testing/Service"/>
</xsl:template>

<xsl:template match="Testing/Service">
	<xsl:choose>
		<xsl:when test="$serviceType = @name">
			<h1 id="headline" class="margin_10_0_0_0"><xsl:value-of select="Headline"/></h1>
			<xsl:choose>
				<xsl:when test="$selectedContent = 'overview'">
					<img src="" border="0" class="serviceLogoPosition">
						<xsl:attribute name="src">../images/nondestructive/<xsl:value-of select="$serviceType"/>Logo.gif</xsl:attribute>
						<xsl:attribute name="alt"><xsl:value-of select="$serviceType" disable-output-escaping="yes" /></xsl:attribute>
					</img>
					<xsl:for-each select="Paragraph">
						<p><xsl:value-of select="." disable-output-escaping="yes" /></p>
					</xsl:for-each>
				</xsl:when>
				<xsl:when test="$selectedContent = 'image_gallery'">
					<xsl:choose>
  						<xsl:when test="$image = ''"> 
							<xsl:apply-templates select="Gallery/Image"/>
						</xsl:when>
   						<xsl:otherwise>
							<img name="xsl" border="0" width="400" class="margin_0_10_10_0">
								<xsl:attribute name="src">../images/nondestructive/<xsl:value-of select="$serviceType"/>/images/<xsl:value-of select="$image"/>.jpg</xsl:attribute>
								<xsl:attribute name="alt"><xsl:value-of select="Gallery/Image/@alt" disable-output-escaping="yes" /></xsl:attribute>
							</img>
							<p><xsl:value-of select="Gallery/Image/Info"/></p>
							<p class="fineprint"><a><xsl:attribute name="href">index.asp?selectedContent=<xsl:value-of select="$selectedContent"/>&amp;XMLFile=<xsl:value-of select="$XMLFile"/>&amp;serviceType=<xsl:value-of select="$serviceType"/></xsl:attribute>Return to Thumbnail View</a></p>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
			</xsl:choose>
		</xsl:when>
	</xsl:choose>
</xsl:template>

<xsl:template match="Gallery/Image">
	<a><xsl:attribute name="href">index.asp?selectedContent=<xsl:value-of select="$selectedContent"/>&amp;XMLFile=<xsl:value-of select="$XMLFile"/>&amp;serviceType=<xsl:value-of select="$serviceType"/>&amp;image=<xsl:value-of select="@alt"/></xsl:attribute>
	<img name="xsl" border="0" width="90" class="margin_0_2_2_0">
		<xsl:attribute name="src">../images/nondestructive/<xsl:value-of select="$serviceType"/>/thumbnails/<xsl:value-of select="@alt"/><xsl:value-of select="File"/></xsl:attribute>
		<xsl:attribute name="alt"><xsl:value-of select="@alt" disable-output-escaping="yes" /></xsl:attribute>
		<xsl:attribute name="title"><xsl:value-of select="Info" disable-output-escaping="yes" /></xsl:attribute>
	</img></a>						
</xsl:template>

</xsl:stylesheet>