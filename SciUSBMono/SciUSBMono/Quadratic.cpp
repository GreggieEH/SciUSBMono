#include "stdafx.h"
#include "Quadratic.h"

CQuadratic::CQuadratic(void) :
	m_minWave(0.0),
	m_maxWave(0.0),
	m_nValues(0),
	m_pCorrection(NULL),
	m_linear(1.0),
	m_offset(0.0),
	m_quadratic(0.0)
{
}

CQuadratic::~CQuadratic(void)
{
	if (NULL != this->m_pCorrection)
	{
		delete [] this->m_pCorrection;
		this->m_pCorrection = NULL;
	}
	this->m_nValues = 0;
}

double CQuadratic::GetQuadFactor()
{
	return this->m_quadratic;
}

void CQuadratic::SetQuadFactor(
							double		quadFactor)
{
	this->m_quadratic = quadFactor;
}

double CQuadratic::GetLinearFactor()
{
	return this->m_linear;
}

void CQuadratic::SetLinearFactor(
							double		linearFactor)
{
	this->m_linear = linearFactor;
}

double CQuadratic::GetOffsetFactor()
{
	return this->m_offset;
}

void CQuadratic::SetOffsetFactor(
							double		offsetFactor)
{
	this->m_offset = offsetFactor;
}

double CQuadratic::GetMinWavelength()
{
	return this->m_minWave;
}

void CQuadratic::SetMinWavelength(
							double		minWavelength)
{
	this->m_minWave = minWavelength;
}

double CQuadratic::GetMaxWavelength()
{
	return this->m_maxWave;
}

void CQuadratic::SetMaxWavelength(
							double		maxWavelength)
{
	this->m_maxWave = maxWavelength;
}

BOOL CQuadratic::MakeCorrections()
{
    long		i;             // index over the correction factors
    long		nWaves;			// number of wavelengths for which the correction factor is calculated
    double		wave;			// wavelength in nanometers
    double		A;	//             ' quadratic equation factor
    double		B;	//            ' quadratic equation factor
    double		C;		//           ' quadratic equation factor
    double		vplus;
    double		vminus;
    
    // set the array dimensions
    nWaves = this->GetMaxWavelength() - this->GetMinWavelength() + 1;
    if ( nWaves > 0)
	{
        // calculate the inverse of the correction factor
        // create the look up table
		if (NULL != this->m_pCorrection)
			delete [] this->m_pCorrection;
		this->m_pCorrection = new double [nWaves];
		this->m_nValues = nWaves;
		if (this->GetQuadFactor() != 0.0)
		{
			for (i=0; i<nWaves; i++)
			{
				wave = this->GetMinWavelength() + (i*1.0);
                C = this->GetOffsetFactor() - wave;             // quadratic equation offset factor
                B = this->GetLinearFactor();					// quadratic equation linear factor
                A = this->GetQuadFactor();
                vplus = (-B + sqrt((B * B) - (4.0 * A * C))) / (2.0 * A);
                this->m_pCorrection[i] = vplus;
			}
		}
		else
		{
			for (i=0; i<nWaves; i++)
			{
				this->m_pCorrection[i] = this->GetMinWavelength() + (i * 1.0);
			}
		}
	}
	return TRUE;
}

double CQuadratic::ExpToActual(
							double		waveIn)
{
    double			qPart;
    double			lPart;
	double			retVal	= waveIn;

    // check the input wavelength
    if (NULL != this->m_pCorrection && waveIn >= this->GetMinWavelength() &&
		waveIn <= this->GetMaxWavelength())
	{
        qPart = this->GetQuadFactor() * waveIn * waveIn;
        lPart = this->GetLinearFactor() * waveIn;
		retVal = qPart + lPart + this->GetOffsetFactor();
	}
	return retVal;
}

double CQuadratic::ActualToExp(
							double		waveIn)
{
	double		waveDiff;          // difference between waveIn and the minimum wavelength
    long		X1;					// lower index
    double		rat;
	double		retVal		= waveIn;

    // check the input wavelength
    if (NULL != this->m_pCorrection && waveIn >= this->GetMinWavelength() && 
		waveIn <= this->GetMaxWavelength())
	{
        // use the look up table to determine the experimental wavelength
        waveDiff = waveIn - this->GetMinWavelength();
        // interpolate to the index
        X1 = (long) floor(waveDiff);
        if (X1 < (this->m_nValues - 1))
		{
			// interpolate to the value
            rat = waveDiff - (1.0 * X1);
			retVal = this->m_pCorrection[X1] + (rat * (this->m_pCorrection[X1+1] - this->m_pCorrection[X1]));
		}
	}
	return retVal;
}
