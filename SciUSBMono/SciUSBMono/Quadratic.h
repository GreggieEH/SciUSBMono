#pragma once

class CQuadratic
{
public:
	CQuadratic(void);
	~CQuadratic(void);
	double				GetQuadFactor();
	void				SetQuadFactor(
							double		quadFactor);
	double				GetLinearFactor();
	void				SetLinearFactor(
							double		linearFactor);
	double				GetOffsetFactor();
	void				SetOffsetFactor(
							double		offsetFactor);
	double				GetMinWavelength();
	void				SetMinWavelength(
							double		minWavelength);
	double				GetMaxWavelength();
	void				SetMaxWavelength(
							double		maxWavelength);
	BOOL				MakeCorrections();
	double				ExpToActual(
							double		waveIn);
	double				ActualToExp(
							double		waveIn);
private:
	double				m_minWave;
	double				m_maxWave;
	long				m_nValues;
	double			*	m_pCorrection;
	double				m_linear;
	double				m_offset;
	double				m_quadratic;
};
