#pragma once

class CStartupValues
{
public:
	CStartupValues(long grating, double position, BOOL autoGrating);
	~CStartupValues(void);
	long			GetGrating();
	double			GetPosition();
	BOOL			GetAutoGrating();
private:
	long			m_grating;
	double			m_position;
	BOOL			m_autoGrating;
};
