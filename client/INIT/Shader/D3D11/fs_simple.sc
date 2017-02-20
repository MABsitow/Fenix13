$input v_color0, v_texcoord0

/*
 * Copyright 2011-2015 Branimir Karadzic. All rights reserved.
 * License: http://www.opensource.org/licenses/BSD-2-Clause
 */

#include "../Common.slib"

SAMPLER2D(u_texColor, 0);

void main()
{
	vec4 pixel = texture2D(u_texColor, v_texcoord0);

	if (pixel.r == 0 && pixel.g == 0 && pixel.b == 0)
		discard;
	else
		gl_FragColor = texture2D(u_texColor, v_texcoord0) * v_color0;
}
