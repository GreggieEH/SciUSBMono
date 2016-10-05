#include "stdafx.h"
#include "StartupValues.h"

CStartupValues::CStartupValues(long grating, double position, BOOL autoGrating) :
	m_grating(grating),
	m_position(0.0),
	m_autoGrating(autoGrating)
{
}

CStartupValues::~CStartupValues(void)
{
}

long CStartupValues::GetGrating()
{
	return this->m_grating;
}

double CStartupValues::GetPosition()
{
	return this->m_position;
}

BOOL CStartupValues::GetAutoGrating()
{
	return this->m_autoGrating;
}
