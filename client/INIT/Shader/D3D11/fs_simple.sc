$input v_color0, v_texcoord0

/*
 * Copyright 2011-2015 Branimir Karadzic. All rights reserved.
 * License: http://www.opensource.org/licenses/BSD-2-Clause
 */

#include "../Common.slib"

SAMPLER2D(u_texColor, 0);

void main()
{
	gl_FragColor = texture2D(u_texColor, v_texcoord0) * v_color0;
}
