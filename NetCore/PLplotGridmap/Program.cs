using PLplot;

class Program
{
    static void Main(string[] args)
    {
        // http://plplot.org/examples.php?demo=02
        using (var pl = new PLStream())
        {
            pl.sdev("png");
            pl.sfnam("imagerect%n.png");

            pl.sfam(1, 1, 10);
            // pl.gfam(out int p_fam, out int p_num, out int p_bmax);

            pl.init();

            demo1(pl);
            demo2(pl);
        }
    }

    static void fillColor(PLStream pl, out int[] r, out int[] g, out int[] b)
    {
        r = new int[116];
        g = new int[116];
        b = new int[116];

        double lmin = 0.15, lmax = 0.85;

        for (int i = 0; i <= 15; i++)
        {
            pl.gcol0(i, out r[i], out g[i], out b[i]);
        }

        for (int i = 0; i <= 99; i++)
        {
            double h, l, s;
            double r1, g1, b1;

            h = (360.0 / 10.0) * (i % 10);
            l = lmin + (lmax - lmin) * (i / 10) / 9.0;
            s = 1.0;

            pl.hlsrgb(h, l, s, out r1, out g1, out b1);
            r[i + 16] = (int)(r1 * 255.001);
            g[i + 16] = (int)(g1 * 255.001);
            b[i + 16] = (int)(b1 * 255.001);
        }
    }

    private static void demo2(PLStream pl)
    {
        pl.bop();
        pl.ssub(10, 10);

        fillColor(pl, out int[] r, out int[] g, out int[] b);
        pl.scmap0(r, g, b);

        draw_gridmap(pl, 100, 16);

        pl.eop();
    }

    private static void demo1(PLStream pl)
    {
        pl.bop();

        int nx = 4;
        int ny = 4;

        {
            pl.ssub(nx, ny);

            draw_gridmap(pl, nx * ny, 0);
        }

        pl.eop();
    }

    private static void draw_gridmap(PLStream pl, int nw, int cmap0_offset)
    {
        double vmin = 0.0, vmax = 1.0;

        for (int i = 0; i < nw; i++)
        {
            pl.col0(i + cmap0_offset);

            pl.adv(0);

            pl.vpor(vmin, vmax, vmin, vmax);
            pl.wind(vmin, vmax, vmin, vmax);

            plfbox(pl);
        }
    }

    private static void plfbox(PLStream pl)
    {
        double[] x = { 0, 0, 1.0, 1.0 };
        double[] y = { 0, 1.0, 1.0, 0 };

        pl.fill(x, y);
    }
}
