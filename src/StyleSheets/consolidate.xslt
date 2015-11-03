<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0"
                xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                >
	<xsl:output method="xml" indent="yes" />

	<xsl:template match="Form[@fileName]">
		<xsl:variable name='fileName' select='@fileName'/>
		<xsl:variable name='contents' select='document($fileName)'/>
		<xsl:comment> Source: <xsl:value-of select ='@fileName'/>
</xsl:comment>
		<xsl:apply-templates select="$contents/Form"/>
	</xsl:template>
	
	<xsl:template match="Report[@fileName]">
		<xsl:variable name='fileName' select='@fileName'/>
		<xsl:variable name='contents' select='document($fileName)'/>
		<xsl:comment> Source: <xsl:value-of select ='@fileName'/>
</xsl:comment>
		<xsl:apply-templates select="$contents/Report"/>
	</xsl:template>

	<xsl:template match="Module[@fileName]">
		<xsl:variable name='fileName' select='@fileName'/>
		<xsl:variable name='contents' select='document($fileName)'/>
		<xsl:comment> Source: <xsl:value-of select ='@fileName'/>
</xsl:comment>
		<xsl:apply-templates select="$contents/Module"/>
	</xsl:template>

	<xsl:template match="*">
		<xsl:copy>
			<xsl:apply-templates select="@*"/>
			<xsl:apply-templates select="node()"/>
		</xsl:copy>
	</xsl:template>

	<xsl:template match="@*">
		<xsl:copy/>
	</xsl:template>
  
</xsl:stylesheet>
